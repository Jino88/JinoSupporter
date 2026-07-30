using System.Diagnostics;
using System.Text;

namespace JinoSupporter.Web.Services;

public sealed class AppActivityLogger(
    IHttpContextAccessor httpContextAccessor,
    NgRateSettingsService ngRateSettings)
{
    private static readonly object FileLock = new();

    public string LogDirectory => Path.Combine(AppStoragePaths.RootDirectory, "logs");

    public string CurrentLogPath => Path.Combine(LogDirectory, $"web-{DateTime.Now:yyyyMMdd}.log");

    public string CurrentWorker
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(ngRateSettings.LoginId))
                return ngRateSettings.LoginId.Trim();

            string? webUser = httpContextAccessor.HttpContext?.User?.Identity?.Name;
            if (!string.IsNullOrWhiteSpace(webUser))
                return webUser.Trim();

            return string.IsNullOrWhiteSpace(Environment.UserName)
                ? "unknown"
                : Environment.UserName;
        }
    }

    public void Log(string area, string action)
    {
        string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {area}] {CurrentWorker} : {action}";
        Debug.WriteLine(line);
        Console.WriteLine(line);
        WriteFileLine(line);
    }

    private void WriteFileLine(string line)
    {
        try
        {
            Directory.CreateDirectory(LogDirectory);
            lock (FileLock)
            {
                File.AppendAllText(CurrentLogPath, line + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch
        {
            // Logging must never take the web process down.
        }
    }
}
