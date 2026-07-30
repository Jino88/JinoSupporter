using System.Diagnostics;

namespace BmesNgRateStandalone.Services;

public sealed class AppActivityLogger(NgRateSettingsService ngRateSettings)
{
    public string CurrentWorker
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(ngRateSettings.LoginId))
                return ngRateSettings.LoginId.Trim();

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
