using System.Text;
using System.Text.Json;
using ExcelExporter;
using WinForms = System.Windows.Forms;

// Web reads stdout as UTF-8 — match here so Korean / non-ASCII paths stay intact.
// On Korean Windows the default Console.OutputEncoding is CP949, which mojibakes
// emitted JSON by the time it reaches the web layer.
Console.OutputEncoding = Encoding.UTF8;
Console.InputEncoding  = Encoding.UTF8;

// office.dll (Microsoft.Office.Core PIA) isn't on nuget. Excel COM transitively
// references it. Probe Office install folders before any Excel COM call so the
// CLR can resolve `office, Version=15.0.0.0` from the user's Office installation.
EnsureOfficeCorePia();

// Two commands:
//   1) pick-folder [--initial <dir>]
//      - Opens a Windows FolderBrowserDialog. Stdout: { "kind": "folder", "path": "<picked>" }
//        On cancel: { "kind": "cancelled" }. Exit 0 on success, 1 on cancel.
//   2) clean --source <dir> --dest <dir>
//      - Recursively scans <source> for *.xlsx / *.xlsm, runs DrmCleanService on each.
//        Stdout: one JSON object per line so the caller can stream parse events.
//
// All output is JSON-Lines on stdout. Diagnostic text (if any) goes to stderr.

if (args.Length == 0)
{
    Console.Error.WriteLine("Usage: pick-folder [--initial <dir>] | clean --source <dir> --dest <dir>");
    return 64;
}

try
{
    return args[0].ToLowerInvariant() switch
    {
        "pick-folder" => RunPickFolder(args),
        "clean"       => RunClean(args),
        _ => Bail($"Unknown command: {args[0]}", 64),
    };
}
catch (Exception ex)
{
    EmitJson(new { kind = "fatal", error = ex.GetType().Name, message = ex.Message });
    Console.Error.WriteLine(ex);
    return 99;
}

static int RunPickFolder(string[] args)
{
    string? initial = ArgValue(args, "--initial");

    // STA required for shell dialogs.
    string? picked = null;
    bool cancelled = false;

    var t = new Thread(() =>
    {
        WinForms.Application.SetHighDpiMode(WinForms.HighDpiMode.SystemAware);

        // Spawn an invisible top-most owner so the dialog renders above the
        // browser (the web page that triggered us) and gets keyboard focus.
        // Without an owner the FolderBrowserDialog can land behind the active
        // window — looks like the dialog never opened.
        using var owner = new WinForms.Form
        {
            FormBorderStyle = WinForms.FormBorderStyle.None,
            StartPosition   = WinForms.FormStartPosition.CenterScreen,
            Size            = new System.Drawing.Size(1, 1),
            ShowInTaskbar   = false,
            TopMost         = true,
            Opacity         = 0,
        };
        owner.Show();
        owner.Activate();
        owner.BringToFront();

        using var dlg = new WinForms.FolderBrowserDialog
        {
            Description            = "Select folder",
            UseDescriptionForTitle = true,
            ShowNewFolderButton    = true,
        };
        if (!string.IsNullOrEmpty(initial) && Directory.Exists(initial))
            dlg.SelectedPath = initial;

        var result = dlg.ShowDialog(owner);
        if (result == WinForms.DialogResult.OK)
            picked = dlg.SelectedPath;
        else
            cancelled = true;

        owner.Close();
    });
    t.SetApartmentState(ApartmentState.STA);
    t.Start();
    t.Join();

    if (cancelled || picked is null)
    {
        EmitJson(new { kind = "cancelled" });
        return 1;
    }
    EmitJson(new { kind = "folder", path = picked });
    return 0;
}

static int RunClean(string[] args)
{
    string? src  = ArgValue(args, "--source");
    string? dest = ArgValue(args, "--dest");
    bool verbose      = HasFlag(args, "--verbose");
    bool keepFormats  = HasFlag(args, "--keep-formats");
    if (string.IsNullOrWhiteSpace(src) || string.IsNullOrWhiteSpace(dest))
        return Bail("clean requires --source <file-or-dir> --dest <dir>", 64);
    if (!Directory.Exists(src) && !File.Exists(src))
        return Bail($"Source not found: {src}", 65);

    // ShapeRenderer captures each shape via System.Windows.Forms.Clipboard,
    // which requires STA. The console app's main thread is MTA by default
    // (top-level statements), so without this jump every Clipboard.GetImage()
    // returns null → no images survive in the converted file.
    int exitCode = 0;
    Exception? failure = null;
    var t = new Thread(() =>
    {
        try { exitCode = RunCleanCore(src!, dest!, verbose, keepFormats); }
        catch (Exception ex) { failure = ex; }
    });
    t.SetApartmentState(ApartmentState.STA);
    t.Start();
    t.Join();
    if (failure != null) throw failure;
    return exitCode;
}

static int RunCleanCore(string src, string dest, bool verbose, bool keepFormats)
{
    Directory.CreateDirectory(dest);
    bool folderMode = !File.Exists(src);

    // dest 폴더 구조:
    //   <dest>/drm_clean/  ← 변환된 _clean.xlsx 가 여기 떨어진다 (DrmCleanService)
    //   <dest>/_done/      ← 변환 성공한 source 원본을 여기로 이동 (folder 모드 only)
    string cleanDir = Path.Combine(dest, "drm_clean");

    // src 가 파일이면 단일 파일 모드, 폴더면 재귀 스캔. Input Data (Test) 페이지가
    // 한 파일 업로드 후 호출하는 케이스에 쓰인다.
    List<string> files;
    int alreadyDone = 0;
    if (!folderMode)
    {
        files = new List<string> { src };
    }
    else
    {
        // Dest 가 source 의 하위 폴더면 (e.g. source\1\drm_clean), recursive 스캔이
        // 이전 실행의 출력 파일을 다시 입력으로 잡아서 *_clean_clean.xlsx 가 생긴다.
        // dest 의 정규화된 절대경로 아래에 있는 모든 파일은 스캔 대상에서 제외.
        string destFull = Path.GetFullPath(dest)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        bool IsUnderDest(string p) =>
            Path.GetFullPath(p).StartsWith(destFull, StringComparison.OrdinalIgnoreCase);

        var allCandidates = Directory.EnumerateFiles(src, "*.*", SearchOption.AllDirectories)
            .Where(p => (p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                      || p.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                     && !Path.GetFileName(p).StartsWith("~$", StringComparison.Ordinal)
                     && !IsUnderDest(p))
            .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Pre-skip: 출력 (`<cleanDir>/<safeName>_clean.xlsx`) 이 이미 있으면
        // 다시 처리하지 않는다. 이전 실행이 중단됐다가 재시작될 때 처음부터
        // 다시 도는 낭비를 막는다.
        files = allCandidates.Where(p =>
        {
            string srcName = Path.GetFileNameWithoutExtension(p);
            string expected = Path.Combine(cleanDir, $"{SafeFileName(srcName)}_clean.xlsx");
            return !File.Exists(expected);
        }).ToList();
        alreadyDone = allCandidates.Count - files.Count;

        if (alreadyDone > 0)
        {
            EmitJson(new {
                kind = "log",
                message = $"[skip] {alreadyDone} 파일은 이미 변환됨 (output 존재) — 새로 처리: {files.Count}"
            });
        }
    }

    EmitJson(new { kind = "scan", count = files.Count, source = src, dest });

    if (files.Count == 0)
    {
        EmitJson(new { kind = "done", ok = 0, fail = 0 });
        return 0;
    }

    var svc = new DrmCleanService(
        log: line => EmitJson(new { kind = "log", message = line }),
        verbose: verbose,
        keepFormats: keepFormats);

    var results = svc.CleanMany(
        files,
        dest,
        onProgress: p => EmitJson(new
        {
            kind   = "progress",
            current = p.Current,
            total  = p.Total,
            file   = Path.GetFileName(p.CurrentSource),
        }));

    // DrmCleanService.Dispose() releases Excel COM and may touch Office.Core
    // types — if office.dll resolution still fails, the cleanup throw shouldn't
    // sink the whole run. Results are already collected at this point.
    try { svc.Dispose(); }
    catch (Exception ex)
    {
        EmitJson(new { kind = "log", message = $"⚠ Excel cleanup warning: {ex.Message}" });
    }

    // 폴더 모드에서만: 성공한 source 파일을 <dest>/_done/ 로 이동.
    // 다음 스캔에서 자동으로 빠져나간다 (원본 위치엔 더 이상 존재 X).
    // 단일 파일 모드 (Input Data Test 의 임시 업로드 등) 는 건너뜀.
    int moved = 0, moveFailed = 0;
    if (folderMode)
    {
        string doneDir = Path.Combine(dest, "_done");
        Directory.CreateDirectory(doneDir);

        foreach (var r in results.Where(x => x.Success))
        {
            try
            {
                string relPath = Path.GetRelativePath(src, r.SourcePath);
                string targetPath = Path.Combine(doneDir, relPath);
                string? targetParent = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrEmpty(targetParent)) Directory.CreateDirectory(targetParent);
                if (File.Exists(targetPath))
                {
                    // _done 에 같은 이름이 이미 있으면 timestamp 붙여 보존
                    string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string tDir = Path.GetDirectoryName(targetPath)!;
                    string tName = Path.GetFileNameWithoutExtension(targetPath);
                    string tExt = Path.GetExtension(targetPath);
                    targetPath = Path.Combine(tDir, $"{tName}.{stamp}{tExt}");
                }
                File.Move(r.SourcePath, targetPath);
                moved++;
            }
            catch (Exception ex)
            {
                moveFailed++;
                EmitJson(new {
                    kind = "log",
                    message = $"⚠ source 이동 실패 [{Path.GetFileName(r.SourcePath)}]: {ex.Message}"
                });
            }
        }
        if (moved > 0 || moveFailed > 0)
        {
            EmitJson(new {
                kind = "log",
                message = $"[archive] {moved} source 파일을 <dest>/_done/ 으로 이동 (실패: {moveFailed})"
            });
        }
    }

    int ok = 0, fail = 0;
    foreach (var r in results)
    {
        if (r.Success)
        {
            ok++;
            EmitJson(new
            {
                kind = "result",
                success = true,
                strategy = r.Strategy.ToString(),
                source = r.SourcePath,
                dest = r.DestPath,
            });
        }
        else
        {
            fail++;
            EmitJson(new
            {
                kind = "result",
                success = false,
                source = r.SourcePath,
                error = r.Error,
            });
        }
    }
    EmitJson(new { kind = "done", ok, fail });
    return fail > 0 ? 2 : 0;
}

static string? ArgValue(string[] args, string flag)
{
    for (int i = 0; i < args.Length - 1; i++)
        if (string.Equals(args[i], flag, StringComparison.Ordinal))
            return args[i + 1];
    return null;
}

// DrmCleanService.SafeName 과 동일 로직 — Pre-skip 단계에서 출력 경로를
// 미리 계산하기 위해 로컬에 복제. 양쪽이 같이 움직여야 한다.
static string SafeFileName(string raw)
{
    var invalid = Path.GetInvalidFileNameChars();
    var chars = raw.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
    var cleaned = new string(chars).Trim(' ', '.');
    return string.IsNullOrEmpty(cleaned) ? "untitled" : cleaned;
}

static bool HasFlag(string[] args, string flag)
{
    foreach (var a in args)
        if (string.Equals(a, flag, StringComparison.Ordinal))
            return true;
    return false;
}

static int Bail(string message, int code)
{
    EmitJson(new { kind = "error", message });
    Console.Error.WriteLine(message);
    return code;
}

static void EmitJson(object payload)
{
    string json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    });
    Console.WriteLine(json);
    Console.Out.Flush();
}

/// <summary>Locate the user's office.dll PIA from a known Office install path
/// and hand it to the CLR via <c>AssemblyResolve</c>. Without this, Excel COM
/// throws <c>FileNotFoundException</c> for <c>office, Version=15.0.0.0</c>
/// because that PIA isn't published on nuget.org.</summary>
static void EnsureOfficeCorePia()
{
    string[] roots =
    {
        Environment.GetEnvironmentVariable("ProgramFiles") ?? @"C:\Program Files",
        Environment.GetEnvironmentVariable("ProgramFiles(x86)") ?? @"C:\Program Files (x86)",
    };
    string[] versions = { "OFFICE16", "OFFICE15", "OFFICE14" };
    string[] subPaths = { "office.dll", @"DCF\office.dll" };

    string? officeDll = null;
    foreach (var root in roots.Distinct())
    {
        foreach (var ver in versions)
        {
            foreach (var sub in subPaths)
            {
                string candidate = Path.Combine(root, "Microsoft Office", "Root", ver, sub);
                if (File.Exists(candidate)) { officeDll = candidate; break; }

                candidate = Path.Combine(root, "Microsoft Office", ver, sub);
                if (File.Exists(candidate)) { officeDll = candidate; break; }

                candidate = Path.Combine(root, "Common Files", "microsoft shared", ver, sub);
                if (File.Exists(candidate)) { officeDll = candidate; break; }
            }
            if (officeDll != null) break;
        }
        if (officeDll != null) break;
    }

    if (officeDll == null) return;   // Office may not be installed — let the COM call fail visibly.

    string resolvedPath = officeDll;
    AppDomain.CurrentDomain.AssemblyResolve += (_, e) =>
    {
        if (e.Name.StartsWith("office,", StringComparison.OrdinalIgnoreCase)
         || e.Name.StartsWith("Microsoft.Office.Core,", StringComparison.OrdinalIgnoreCase))
        {
            try { return System.Reflection.Assembly.LoadFrom(resolvedPath); }
            catch { }
        }
        return null;
    };
}
