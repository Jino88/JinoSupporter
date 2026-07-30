using Microsoft.Extensions.Logging;

namespace BmesNgRateStandalone;

public static class StandaloneErrorLog
{
    private static readonly object Sync = new();

    public static string LogDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "JinoWorkHost",
        "logs");

    public static string LogPath { get; } = Path.Combine(LogDirectory, "bmes-ngrate-standalone.log");

    public static void Write(string area, string message)
    {
        try
        {
            Directory.CreateDirectory(LogDirectory);
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{area}] {message}";
            lock (Sync)
            {
                File.AppendAllText(LogPath, line + Environment.NewLine);
            }
        }
        catch
        {
            // Logging must never crash the desktop host.
        }
    }

    public static void Write(string area, Exception exception)
    {
        try
        {
            Directory.CreateDirectory(LogDirectory);
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{area}]{Environment.NewLine}{exception}{Environment.NewLine}";
            lock (Sync)
            {
                File.AppendAllText(LogPath, line);
            }
        }
        catch
        {
            // Logging must never crash the desktop host.
        }
    }
}

public sealed class StandaloneFileLoggerProvider : ILoggerProvider
{
    public ILogger CreateLogger(string categoryName) => new StandaloneFileLogger(categoryName);

    public void Dispose()
    {
    }

    private sealed class StandaloneFileLogger(string categoryName) : ILogger
    {
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

        public bool IsEnabled(LogLevel logLevel) => logLevel >= LogLevel.Warning;

        public void Log<TState>(
            LogLevel logLevel,
            EventId eventId,
            TState state,
            Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel)) return;

            string message = formatter(state, exception);
            if (exception is not null)
                StandaloneErrorLog.Write($"{logLevel} {categoryName}", exception);
            else if (!string.IsNullOrWhiteSpace(message))
                StandaloneErrorLog.Write($"{logLevel} {categoryName}", message);
        }
    }
}
