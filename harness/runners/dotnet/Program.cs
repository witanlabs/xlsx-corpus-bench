// EPPlus / ClosedXML runner.
//
// Per file, measures:
//   load      — open the workbook and touch every sheet's dimension
//   roundtrip — save to out dir, reopen the saved copy with the same library
//   recalc    — record cached values of all formula cells, run the library's
//               calculation engine, compare computed values to cached
//
// Emits one JSON line per file to --out (append mode, resumable).

using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using OfficeOpenXml;

const int PerFileTimeoutMs = 120_000;

var opts = new Dictionary<string, string>();
for (var i = 0; i + 1 < args.Length; i += 2)
{
    if (args[i].StartsWith("--")) opts[args[i][2..]] = args[i + 1];
}
string lib = opts["lib"], corpus = opts["corpus"], manifest = opts["manifest"],
    outFile = opts["out"], outDir = opts["out-dir"];
var mode = opts.GetValueOrDefault("mode", "bench"); // bench | recalc-save

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
Directory.CreateDirectory(outDir);

var done = new HashSet<string>();
if (File.Exists(outFile))
{
    foreach (var line in File.ReadLines(outFile))
    {
        if (string.IsNullOrWhiteSpace(line)) continue;
        try { done.Add(JsonDocument.Parse(line).RootElement.GetProperty("sha256").GetString()!); }
        catch { /* partial line from a previous crash */ }
    }
}

using var output = new StreamWriter(outFile, append: true);
foreach (var line in File.ReadLines(manifest))
{
    if (string.IsNullOrWhiteSpace(line)) continue;
    var rec = JsonDocument.Parse(line).RootElement;
    var sha = rec.GetProperty("sha256").GetString()!;
    if (done.Contains(sha)) continue;
    var relPath = rec.GetProperty("path").GetString()!;
    var ext = rec.GetProperty("ext").GetString()!;

    var task = mode switch
    {
        "recalc-save" => Task.Run(() => RecalcSave(lib, Path.Combine(corpus, relPath), sha, relPath)),
        "recalc-emit" => Task.Run(() => RecalcEmit(lib, Path.Combine(corpus, relPath), sha, relPath)),
        _ => Task.Run(() => ProcessFile(lib, Path.Combine(corpus, relPath), sha, relPath, ext)),
    };
    if (task.Wait(PerFileTimeoutMs))
    {
        output.WriteLine(JsonSerializer.Serialize(task.Result));
        output.Flush();
    }
    else
    {
        // A wedged calc engine can't be cancelled — its abandoned thread keeps
        // spinning and drags every later file. Record the timeout and exit 3;
        // the orchestrator relaunches a clean process which resumes after this
        // file (its sha is now in the output).
        var result = new
        {
            sha256 = sha, path = relPath, lib,
            load = new { ok = false, ms = (int?)null, error = "timeout" },
            roundtrip = new { ok = false, ms = (int?)null, error = (string?)null, @out = (string?)null },
            recalc = new { supported = true, ok = false, error = (string?)null },
        };
        output.WriteLine(JsonSerializer.Serialize(result));
        output.Flush();
        Environment.Exit(3);
    }
}
return;

// recalc-save mode: load, run the calculation engine, save the recalculated
// workbook to out-dir/<sha>.xlsx so the harness can judge its cached values
// against the Excel-recomputed truth corpus with one shared comparator.
object RecalcSave(string lib, string file, string sha, string relPath)
{
    var outPath = Path.Combine(outDir, sha + ".xlsx");
    try
    {
        if (lib == "epplus")
        {
            using var pkg = new ExcelPackage(new FileInfo(file));
            pkg.Workbook.Calculate();
            pkg.SaveAs(new FileInfo(outPath));
        }
        else
        {
            using var wb = new XLWorkbook(file);
            wb.RecalculateAllFormulas();
            wb.SaveAs(outPath);
        }
        return new { sha256 = sha, path = relPath, lib, ok = true, error = (string?)null };
    }
    catch (Exception e)
    {
        return new { sha256 = sha, path = relPath, lib, ok = false, error = Trim(e) };
    }
}

// recalc-emit mode: run the calculation engine and emit every formula cell's
// computed value as canonical JSON (kind: n/s/b/e like the harness extractor)
// so computation is judged independently of whether SaveAs persists cached
// values (ClosedXML's doesn't, EPPlus's only partially).
object RecalcEmit(string lib, string file, string sha, string relPath)
{
    var cells = new Dictionary<string, object?[]>();
    try
    {
        if (lib == "epplus")
        {
            using var pkg = new ExcelPackage(new FileInfo(file));
            pkg.Workbook.Calculate();
            foreach (var ws in pkg.Workbook.Worksheets)
            {
                if (ws.Dimension == null) continue;
                foreach (var cell in ws.Cells[ws.Dimension.Address])
                {
                    if (string.IsNullOrEmpty(cell.Formula)) continue;
                    cells[$"{ws.Name}!{cell.Address}"] = CanonJson(cell.Value);
                }
            }
        }
        else
        {
            using var wb = new XLWorkbook(file);
            wb.RecalculateAllFormulas();
            foreach (var ws in wb.Worksheets)
            foreach (var cell in ws.CellsUsed(c => c.HasFormula))
            {
                object? v;
                try { v = XlValue(cell.Value); }
                catch (Exception) { v = XLError.IncompatibleValue; }
                cells[$"{ws.Name}!{cell.Address}"] = CanonJson(v);
            }
        }
        return new { sha256 = sha, path = relPath, lib, ok = true, error = (string?)null, cells };
    }
    catch (Exception e)
    {
        return new { sha256 = sha, path = relPath, lib, ok = false, error = Trim(e), cells };
    }
}

// canonical [kind, value] matching harness/cached_values.py: n=number,
// s=string, b=bool(0/1 as number), e=error string
static object?[] CanonJson(object? v) => v switch
{
    null => new object?[] { "s", null },
    bool b => new object?[] { "b", b ? 1.0 : 0.0 },
    DateTime dt => new object?[] { "n", dt.ToOADate() },
    TimeSpan ts => new object?[] { "n", ts.TotalDays },
    ExcelErrorValue ev => new object?[] { "e", ev.ToString() },
    XLError xe => new object?[] { "e", XlErrorString(xe) },
    sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal
        => new object?[] { "n", Convert.ToDouble(v, CultureInfo.InvariantCulture) },
    string s => new object?[] { "s", s },
    _ => new object?[] { "s", v.ToString() },
};

static string XlErrorString(XLError e) => e switch
{
    XLError.CellReference => "#REF!",
    XLError.DivisionByZero => "#DIV/0!",
    XLError.IncompatibleValue => "#VALUE!",
    XLError.NameNotRecognized => "#NAME?",
    XLError.NoValueAvailable => "#N/A",
    XLError.NullValue => "#NULL!",
    XLError.NumberInvalid => "#NUM!",
    _ => "#VALUE!",
};

object ProcessFile(string lib, string file, string sha, string relPath, string ext)
{
    var load = new Dictionary<string, object?> { ["ok"] = false, ["ms"] = null, ["error"] = null };
    var roundtrip = new Dictionary<string, object?> { ["ok"] = false, ["ms"] = null, ["error"] = null, ["out"] = null };
    var recalc = new Dictionary<string, object?> { ["supported"] = true, ["ok"] = false, ["error"] = null };
    var outPath = Path.Combine(outDir, sha + ext);

    var sw = Stopwatch.StartNew();
    try
    {
        if (lib == "epplus") EpplusRun(file, outPath, load, roundtrip, recalc, sw);
        else ClosedXmlRun(file, outPath, load, roundtrip, recalc, sw);
    }
    catch (Exception e)
    {
        var target = (bool)(load["ok"] ?? false) ? ((bool?)roundtrip["ok"] == true ? recalc : roundtrip) : load;
        target["error"] = Trim(e);
        if (target["ms"] == null) target["ms"] = (int)sw.ElapsedMilliseconds;
    }
    return new { sha256 = sha, path = relPath, lib, load, roundtrip, recalc };
}

void EpplusRun(string file, string outPath, Dictionary<string, object?> load,
    Dictionary<string, object?> roundtrip, Dictionary<string, object?> recalc, Stopwatch sw)
{
    using (var pkg = new ExcelPackage(new FileInfo(file)))
    {
        foreach (var ws in pkg.Workbook.Worksheets) _ = ws.Dimension?.Address;
        load["ok"] = true; load["ms"] = (int)sw.ElapsedMilliseconds;

        sw.Restart();
        try
        {
            pkg.SaveAs(new FileInfo(outPath));
            using var reopened = new ExcelPackage(new FileInfo(outPath));
            foreach (var ws in reopened.Workbook.Worksheets) _ = ws.Dimension?.Address;
            roundtrip["ok"] = true; roundtrip["out"] = outPath;
        }
        catch (Exception e)
        {
            roundtrip["error"] = Trim(e);
            if (File.Exists(outPath)) roundtrip["out"] = outPath;
        }
        roundtrip["ms"] = (int)sw.ElapsedMilliseconds;
    }

    sw.Restart();
    try
    {
        // fresh load so recalc compares against pristine cached values
        using var pkg2 = new ExcelPackage(new FileInfo(file));
        var cached = new Dictionary<string, object?>();
        foreach (var ws in pkg2.Workbook.Worksheets)
        {
            if (ws.Dimension == null) continue;
            foreach (var cell in ws.Cells[ws.Dimension.Address])
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                    cached[$"{ws.Name}!{cell.Address}"] = cell.Value;
            }
        }
        pkg2.Workbook.Calculate();
        int mismatches = 0, errors = 0;
        foreach (var ws in pkg2.Workbook.Worksheets)
        {
            if (ws.Dimension == null) continue;
            foreach (var cell in ws.Cells[ws.Dimension.Address])
            {
                var key = $"{ws.Name}!{cell.Address}";
                if (!cached.TryGetValue(key, out var before)) continue;
                var after = cell.Value;
                if (after is ExcelErrorValue) errors++;
                if (!ValuesMatch(before, after)) mismatches++;
            }
        }
        recalc["ok"] = true;
        recalc["formula_cells"] = cached.Count;
        recalc["mismatches"] = mismatches;
        recalc["errors"] = errors;
    }
    catch (Exception e) { recalc["error"] = Trim(e); }
    recalc["ms"] = (int)sw.ElapsedMilliseconds;
}

void ClosedXmlRun(string file, string outPath, Dictionary<string, object?> load,
    Dictionary<string, object?> roundtrip, Dictionary<string, object?> recalc, Stopwatch sw)
{
    using (var wb = new XLWorkbook(file))
    {
        foreach (var ws in wb.Worksheets) _ = ws.RangeUsed()?.RangeAddress;
        load["ok"] = true; load["ms"] = (int)sw.ElapsedMilliseconds;

        sw.Restart();
        try
        {
            wb.SaveAs(outPath);
            using var reopened = new XLWorkbook(outPath);
            foreach (var ws in reopened.Worksheets) _ = ws.RangeUsed()?.RangeAddress;
            roundtrip["ok"] = true; roundtrip["out"] = outPath;
        }
        catch (Exception e)
        {
            roundtrip["error"] = Trim(e);
            if (File.Exists(outPath)) roundtrip["out"] = outPath;
        }
        roundtrip["ms"] = (int)sw.ElapsedMilliseconds;
    }

    sw.Restart();
    try
    {
        using var wb2 = new XLWorkbook(file);
        var cells = new List<(IXLCell cell, object? cached)>();
        foreach (var ws in wb2.Worksheets)
        foreach (var cell in ws.CellsUsed(c => c.HasFormula))
            cells.Add((cell, XlValue(cell.CachedValue)));
        wb2.RecalculateAllFormulas();
        int mismatches = 0, errors = 0;
        foreach (var (cell, before) in cells)
        {
            object? after;
            try { after = XlValue(cell.Value); }
            catch (Exception) { errors++; mismatches++; continue; }
            if (after is XLError) errors++;
            if (!ValuesMatch(before, after)) mismatches++;
        }
        recalc["ok"] = true;
        recalc["formula_cells"] = cells.Count;
        recalc["mismatches"] = mismatches;
        recalc["errors"] = errors;
    }
    catch (Exception e) { recalc["error"] = Trim(e); }
    recalc["ms"] = (int)sw.ElapsedMilliseconds;
}

static object? XlValue(XLCellValue v) => v.Type switch
{
    XLDataType.Blank => null,
    XLDataType.Boolean => v.GetBoolean(),
    XLDataType.Number => v.GetNumber(),
    XLDataType.Text => v.GetText(),
    XLDataType.Error => v.GetError(),
    XLDataType.DateTime => v.GetDateTime().ToOADate(),
    XLDataType.TimeSpan => v.GetTimeSpan().TotalDays,
    _ => v.ToString(),
};

static bool ValuesMatch(object? a, object? b)
{
    a = Canon(a); b = Canon(b);
    if (a == null && b == null) return true;
    if (a is double da && b is double db)
    {
        if (double.IsNaN(da) || double.IsNaN(db)) return double.IsNaN(da) == double.IsNaN(db);
        var tol = Math.Max(1e-9, Math.Abs(da) * 1e-9);
        return Math.Abs(da - db) <= tol;
    }
    return string.Equals(
        Convert.ToString(a, CultureInfo.InvariantCulture),
        Convert.ToString(b, CultureInfo.InvariantCulture),
        StringComparison.Ordinal);
}

static object? Canon(object? v) => v switch
{
    null => null,
    string s when s.Length == 0 => null,
    bool => v,
    DateTime dt => dt.ToOADate(),
    TimeSpan ts => ts.TotalDays,
    sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal
        => Convert.ToDouble(v, CultureInfo.InvariantCulture),
    _ => v.ToString(),
};

static string Trim(Exception e)
{
    var s = $"{e.GetType().Name}: {e.Message}";
    return s.Length > 500 ? s[..500] : s;
}
