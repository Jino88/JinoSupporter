using System.Diagnostics;

namespace JinoSupporter.Web.Services;

public sealed class AppActivityLogger(
    IHttpContextAccessor httpContextAccessor,
    NgRateSettingsService ngRateSettings)
{
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
        string line = $"[{area} {DateTime.Now:HH:mm:ss.fff}] {CurrentWorker} : {action}";
        Debug.WriteLine(line);
        Console.WriteLine(line);
    }
}
