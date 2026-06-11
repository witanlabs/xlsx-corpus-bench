#!/usr/bin/env node
// witan runner — drives the locally built xlsx-serve binary
// (../witan-alfred/bin/publish/xlsx-serve, override with WITAN_XLSX_SERVE).
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
import { fileURLToPath } from 'node:url';
import readline from 'node:readline';

const PER_OP_TIMEOUT_MS = 120_000;
const CONCURRENCY = Number(process.env.WITAN_BENCH_CONCURRENCY || 8);

const args = Object.fromEntries(
  process.argv.slice(2).map((a, i, all) => (a.startsWith('--') ? [a.slice(2), all[i + 1]] : null)).filter(Boolean),
);
const { corpus, manifest, out, 'out-dir': outDir } = args;
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const xlsxServe =
  process.env.WITAN_XLSX_SERVE || path.resolve(repoRoot, '..', 'witan-alfred', 'bin', 'publish', 'xlsx-serve');

if (!fs.existsSync(xlsxServe)) {
  console.error(`xlsx-serve binary not found at ${xlsxServe}`);
  process.exit(1);
}
fs.mkdirSync(outDir, { recursive: true });

// run xlsx-serve, returning {ok, json, error}; non-zero exit with parseable
// stdout is not a process failure (calc exits 2 on errors/changed values)
function run(cliArgs) {
  return new Promise((resolve) => {
    execFile(
      xlsxServe,
      cliArgs,
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
  const r = await run(['--json', 'exec', file, '--expr', 'await xlsx.listSheets(wb)']);
  if (!r.ok) return { ok: false, error: r.error };
  if (r.json.ok === false) return { ok: false, error: (r.json.error?.message ?? 'exec failed').slice(0, 500) };
  return { ok: true, error: null };
}

// save-only rewrite via the JSON-RPC server mode: open + save, one response
// line per request, every line must be ok
function rpcOpenSave(file) {
  return new Promise((resolve) => {
    const proc = execFile(
      xlsxServe,
      [],
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
        if (responses.length >= 2 && responses.every((r) => r.ok)) {
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
    proc.stdin.write(
      JSON.stringify({ id: '1', workbook: file, op: 'open', args: {} }) +
        '\n' +
        JSON.stringify({ id: '2', workbook: file, op: 'save', args: {} }) +
        '\n',
    );
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
  const verify = await run(['--json', 'calc', file, '--verify']);
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
