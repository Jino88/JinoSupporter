using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelDrm
{
    public enum ConvertMode
    {
        Fast,
        Balanced,
        Precise,
        Clipboard,
    }

    public sealed class ConvertResult
    {
        public bool Success { get; init; }
        public string Input { get; init; } = "";
        public string? Output { get; init; }
        public string? Error { get; init; }
        public double ElapsedSeconds { get; init; }
        public int ExitCode { get; init; }
        public string StdErr { get; init; } = "";
    }

    /// <summary>
    /// DRM 보호 xlsx 를 Python CLI(excel_drm_cli.py) 통해 정상 xlsx 로 변환.
    /// 한 번 호출 = 파일 1개. 내부적으로 python.exe subprocess 실행.
    /// </summary>
    public sealed class ExcelDrmCleaner
    {
        public string PythonExe { get; }
        public string ScriptPath { get; }
        public ConvertMode DefaultMode { get; set; } = ConvertMode.Clipboard;

        /// <param name="pythonExe">python.exe 절대 경로 (예: C:\Python311\python.exe)</param>
        /// <param name="scriptPath">excel_drm_cli.py 절대 경로</param>
        public ExcelDrmCleaner(string pythonExe, string scriptPath)
        {
            if (string.IsNullOrWhiteSpace(pythonExe))
                throw new ArgumentException("pythonExe 비어있음", nameof(pythonExe));
            if (string.IsNullOrWhiteSpace(scriptPath))
                throw new ArgumentException("scriptPath 비어있음", nameof(scriptPath));
            if (!File.Exists(scriptPath))
                throw new FileNotFoundException("CLI 스크립트를 찾을 수 없음", scriptPath);

            PythonExe = pythonExe;
            ScriptPath = scriptPath;
        }

        public ConvertResult Convert(string inputPath, string outputPath, ConvertMode? mode = null)
            => ConvertAsync(inputPath, outputPath, mode, CancellationToken.None).GetAwaiter().GetResult();

        public async Task<ConvertResult> ConvertAsync(
            string inputPath,
            string outputPath,
            ConvertMode? mode = null,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(inputPath))
                throw new ArgumentException("inputPath 비어있음", nameof(inputPath));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("outputPath 비어있음", nameof(outputPath));

            var modeStr = (mode ?? DefaultMode) switch
            {
                ConvertMode.Fast => "fast",
                ConvertMode.Balanced => "balanced",
                ConvertMode.Precise => "precise",
                ConvertMode.Clipboard => "clipboard",
                _ => "clipboard",
            };

            var psi = new ProcessStartInfo
            {
                FileName = PythonExe,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                WorkingDirectory = Path.GetDirectoryName(ScriptPath) ?? Environment.CurrentDirectory,
            };
            psi.ArgumentList.Add(ScriptPath);
            psi.ArgumentList.Add("--input");
            psi.ArgumentList.Add(inputPath);
            psi.ArgumentList.Add("--output");
            psi.ArgumentList.Add(outputPath);
            psi.ArgumentList.Add("--mode");
            psi.ArgumentList.Add(modeStr);
            // Python 출력 UTF-8 강제 — 한글 로그 깨짐 방지.
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            psi.EnvironmentVariables["PYTHONUTF8"] = "1";

            using var proc = new Process { StartInfo = psi };
            var stdoutSb = new StringBuilder();
            var stderrSb = new StringBuilder();
            proc.OutputDataReceived += (_, e) => { if (e.Data != null) stdoutSb.AppendLine(e.Data); };
            proc.ErrorDataReceived += (_, e) => { if (e.Data != null) stderrSb.AppendLine(e.Data); };

            if (!proc.Start())
                throw new InvalidOperationException("python.exe 실행 실패");
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();

            using (cancellationToken.Register(() =>
            {
                try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); }
                catch { /* ignore */ }
            }))
            {
                await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
            }

            var stdout = stdoutSb.ToString();
            var stderr = stderrSb.ToString();
            return ParseResult(stdout, stderr, proc.ExitCode, inputPath);
        }

        private static ConvertResult ParseResult(string stdout, string stderr, int exitCode, string inputPath)
        {
            // CLI 는 stdout 마지막 줄에 JSON 한 줄을 출력. 빈 줄 무시하고 끝에서 첫 비어있지 않은 줄 사용.
            string? jsonLine = null;
            var lines = stdout.Split('\n');
            for (int i = lines.Length - 1; i >= 0; i--)
            {
                var t = lines[i].Trim();
                if (t.Length > 0) { jsonLine = t; break; }
            }

            if (jsonLine == null)
            {
                return new ConvertResult
                {
                    Success = false,
                    Input = inputPath,
                    Error = $"CLI 응답 없음 (exit={exitCode}). stderr 확인 필요.",
                    ExitCode = exitCode,
                    StdErr = stderr,
                };
            }

            try
            {
                using var doc = JsonDocument.Parse(jsonLine);
                var root = doc.RootElement;
                var status = root.TryGetProperty("status", out var s) ? s.GetString() : null;
                var input = root.TryGetProperty("input", out var i) ? i.GetString() ?? inputPath : inputPath;
                var output = root.TryGetProperty("output", out var o) ? o.GetString() : null;
                var error = root.TryGetProperty("error", out var e) ? e.GetString() : null;
                var elapsed = root.TryGetProperty("elapsed", out var el) && el.TryGetDouble(out var d) ? d : 0.0;

                return new ConvertResult
                {
                    Success = status == "ok" && exitCode == 0,
                    Input = input,
                    Output = output,
                    Error = error,
                    ElapsedSeconds = elapsed,
                    ExitCode = exitCode,
                    StdErr = stderr,
                };
            }
            catch (JsonException ex)
            {
                return new ConvertResult
                {
                    Success = false,
                    Input = inputPath,
                    Error = $"CLI JSON 파싱 실패: {ex.Message} / 응답: {jsonLine}",
                    ExitCode = exitCode,
                    StdErr = stderr,
                };
            }
        }
    }
}
