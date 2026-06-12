#!/usr/bin/env node
// witan runner — drives the public `witan` CLI (npm i -g witan) by default;
// set WITAN_XLSX_SERVE to a local engine build to run offline (no rate
// limits; what the published numbers use). Public-CLI corpus runs need
// authentication first (`witan auth login`) — anonymous traffic is
// rate-limited at corpus scale.
//
// Per file, measures:
//   load      — `exec <file> --expr 'await xlsx.listSheets(wb)'`
//   roundtrip — copy to out dir, RPC open+save on the copy (stdin JSON-RPC:
//               {"op":"open"} then {"op":"save"} — save always rewrites the
//               file, no recalculation involved, same load→save→reload
//               semantics as every other library), then reload via exec
//   recalc    — `calc <original> --verify --json` (non-mutating);
//               formula_cells = |touched|, mismatches = |changed|
//
// Emits one JSON line per file to --out (append mode, resumable).

import { execFile } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import readline from 'node:readline';

const PER_OP_TIMEOUT_MS = 120_000;
// public CLI goes over the network: default to gentler parallelism there
const CONCURRENCY = Number(
  process.env.WITAN_BENCH_CONCURRENCY || (process.env.WITAN_XLSX_SERVE ? 8 : 2),
);

const args = Object.fromEntries(
  process.argv.slice(2).map((a, i, all) => (a.startsWith('--') ? [a.slice(2), all[i + 1]] : null)).filter(Boolean),
);
const { corpus, manifest, out, 'out-dir': outDir } = args;
// default: the public `witan` CLI (npm i -g witan). WITAN_XLSX_SERVE points
// at a local engine build instead (no network, no rate limits).
const localEngine = process.env.WITAN_XLSX_SERVE || null;
if (localEngine && !fs.existsSync(localEngine)) {
  console.error(`WITAN_XLSX_SERVE set but not found: ${localEngine}`);
  process.exit(1);
}
const BIN = localEngine || 'witan';
const SUB = localEngine ? [] : ['xlsx']; // public CLI namespaces under `xlsx`
fs.mkdirSync(outDir, { recursive: true });

// run xlsx-serve, returning {ok, json, error}; non-zero exit with parseable
// stdout is not a process failure (calc exits 2 on errors/changed values)
function run(cliArgs) {
  return new Promise((resolve) => {
    execFile(
      BIN,
      [...SUB, ...cliArgs],
      { timeout: PER_OP_TIMEOUT_MS, maxBuffer: 256 * 1024 * 1024 },
      (err, stdout, stderr) => {
        try {
          resolve({ ok: true, json: JSON.parse(stdout), error: null });
        } catch {
          const detail = err?.killed
            ? 'timeout'
            : `${String(stderr ?? '').trim() || String(stdout ?? '').trim() || String(err?.message ?? 'no output')}`;
          resolve({ ok: false, json: null, error: detail.slice(0, 500) });
        }
      },
    );
  });
}

async function listSheets(file) {
  const r = await run(['exec', file, '--expr', 'await xlsx.listSheets(wb)', '--json']);
  if (!r.ok) return { ok: false, error: r.error };
  if (r.json.ok === false) return { ok: false, error: (r.json.error?.message ?? 'exec failed').slice(0, 500) };
  return { ok: true, error: null };
}

// save-only rewrite via the JSON-RPC server mode: every response line must
// be ok. Local engine: bare server, requests carry a workbook field.
// Public CLI: `witan xlsx rpc <file>` owns the session — no workbook field,
// no explicit open.
function rpcOpenSave(file) {
  return new Promise((resolve) => {
    const proc = execFile(
      BIN,
      localEngine ? [] : ['xlsx', 'rpc', file],
      { timeout: PER_OP_TIMEOUT_MS, maxBuffer: 64 * 1024 * 1024 },
      (err, stdout, stderr) => {
        const responses = String(stdout ?? '')
          .split('\n')
          .filter((l) => l.trim())
          .map((l) => {
            try {
              return JSON.parse(l);
            } catch {
              return null;
            }
          })
          .filter(Boolean);
        const expected = localEngine ? 2 : 1;
        if (responses.length >= expected && responses.every((r) => r.ok)) {
          resolve({ ok: true, error: null });
        } else {
          const bad = responses.find((r) => !r.ok);
          const detail = err?.killed
            ? 'timeout'
            : bad?.message ?? String(stderr ?? '').trim() ?? 'incomplete rpc response';
          resolve({ ok: false, error: String(detail).slice(0, 500) });
        }
      },
    );
    const requests = localEngine
      ? [
          { id: '1', workbook: file, op: 'open', args: {} },
          { id: '2', workbook: file, op: 'save', args: {} },
        ]
      : [{ id: '1', op: 'save', args: {} }];
    proc.stdin.write(requests.map((r) => JSON.stringify(r)).join('\n') + '\n');
    proc.stdin.end();
  });
}

async function processFile(rec) {
  const file = path.join(corpus, rec.path);
  const result = {
    sha256: rec.sha256,
    path: rec.path,
    lib: 'witan',
    load: { ok: false, ms: null, error: null },
    roundtrip: { ok: false, ms: null, error: null, out: null },
    recalc: { supported: true, ok: false, error: null },
  };

  let t0 = Date.now();
  const loaded = await listSheets(file);
  result.load = { ok: loaded.ok, ms: Date.now() - t0, error: loaded.error };
  if (!loaded.ok) return result;

  const outPath = path.join(outDir, rec.sha256 + rec.ext);
  t0 = Date.now();
  fs.copyFileSync(file, outPath);
  const saved = await rpcOpenSave(outPath);
  if (!saved.ok) {
    result.roundtrip = { ok: false, ms: Date.now() - t0, error: saved.error, out: outPath };
  } else {
    const reopened = await listSheets(outPath);
    result.roundtrip = { ok: reopened.ok, ms: Date.now() - t0, error: reopened.error, out: outPath };
  }

  t0 = Date.now();
  const verify = await run(['calc', file, '--verify', '--json']);
  if (verify.ok && verify.json.touched !== undefined) {
    result.recalc = {
      supported: true,
      ok: true,
      error: null,
      ms: Date.now() - t0,
      formula_cells: Object.keys(verify.json.touched ?? {}).length,
      mismatches: (verify.json.changed ?? []).length,
      errors: (verify.json.errors ?? []).length,
    };
  } else {
    result.recalc = {
      supported: true,
      ok: false,
      ms: Date.now() - t0,
      error: verify.error ?? 'unexpected calc output shape',
    };
  }
  return result;
}

const done = new Set();
if (fs.existsSync(out)) {
  for (const line of fs.readFileSync(out, 'utf8').split('\n')) {
    if (!line.trim()) continue;
    try {
      done.add(JSON.parse(line).sha256);
    } catch {}
  }
}

const pending = [];
const rl = readline.createInterface({ input: fs.createReadStream(manifest) });
for await (const line of rl) {
  if (!line.trim()) continue;
  const rec = JSON.parse(line);
  if (!done.has(rec.sha256)) pending.push(rec);
}

const outStream = fs.createWriteStream(out, { flags: 'a' });
let next = 0;
let completed = 0;
async function worker() {
  while (next < pending.length) {
    const rec = pending[next++];
    const res = await processFile(rec);
    outStream.write(JSON.stringify(res) + '\n');
    completed += 1;
    if (completed % 250 === 0) console.error(`witan: ${completed}/${pending.length}`);
  }
}
await Promise.all(Array.from({ length: CONCURRENCY }, worker));
outStream.end();
