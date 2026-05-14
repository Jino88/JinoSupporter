# ExcelDrmCli

DRM-locked `.xlsx` → clean `.xlsx` converter. Replaces the previous
`External/ExcelExporter` C# submodule. C# callers invoke the Python CLI
through `ExcelDrmCleaner.cs` (subprocess + JSON line on stdout).

## Files

| File | Role |
|---|---|
| `excel_drm_cli.py`     | CLI entry. `--input` / `--output` / `--mode`. Emits one JSON line on stdout. |
| `excel_drm_clean.py`   | Conversion engine (imported by the CLI). |
| `ExcelDrmCleaner.cs`   | .NET 8 wrapper. Spawns `python.exe`, parses the JSON line, returns `ConvertResult`. |
| `requirements.txt`     | `pywin32`, `openpyxl`, `pillow`. |

## Prereqs

- Python 3.10+
- `pip install -r requirements.txt`

## CLI usage

```
python excel_drm_cli.py --input <원본.xlsx> --output <대상.xlsx> --mode clipboard
```

- `--mode`: `fast` | `balanced` | `precise` | `clipboard` (default `clipboard`)
- stdout: progress is on **stderr**; stdout's last non-empty line is JSON:
  - success: `{"status":"ok","input":"...","output":"...","elapsed":12.3}`
  - failure: `{"status":"fail","input":"...","error":"...","elapsed":0.5}`
- exit code: `0` = success, `1` = conversion failure, `2` = arg error.

## C# usage

```csharp
var cleaner = new ExcelDrm.ExcelDrmCleaner(
    pythonExe:  @"C:\Python311\python.exe",   // or .venv\Scripts\python.exe
    scriptPath: @"...\External\ExcelDrmCli\excel_drm_cli.py");

var result = await cleaner.ConvertAsync(input, output, ExcelDrm.ConvertMode.Clipboard, ct);
if (!result.Success) throw new InvalidOperationException(result.Error);
```

`ExcelDrmCleaner.cs` is compiled into `JinoSupporter.Web` via a linked
`<Compile Include>` item in `JinoSupporter.Web.csproj`.
