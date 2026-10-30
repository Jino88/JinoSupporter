namespace PortalGNLAB.Web.Services;

/// <summary>
/// Runs the one-shot history cache fill (NG Rate, then F-COST RAW) from a start date up to
/// <see cref="HistoryCachePolicy.LastHistoryDay"/> on a background task, so it keeps going when the
/// Setting page is closed. One run at a time; already settled days are skipped, so a run can be
/// stopped and started again.
/// </summary>
public sealed class HistoryCacheJobService(IServiceScopeFactory scopes, AppActivityLogger activity)
{
    private const int MaxLogLines = 300;
    // F-COST RAW is one API call per day; keep a short gap so a multi-year fill stays gentle on BMES.
    private const int FCostDelayMs = 300;

    private readonly object _gate = new();
    private readonly List<string> _log = [];
    private CancellationTokenSource? _cts;

    public bool IsRunning { get; private set; }
    public DateTime? StartedAt { get; private set; }
    public DateTime? FinishedAt { get; private set; }
    public string Status { get; private set; } = string.Empty;

    /// <summary>Raised on every log line and state change (from a background thread).</summary>
    public event Action? Changed;

    public IReadOnlyList<string> Log
    {
        get { lock (_gate) return _log.ToList(); }
    }

    /// <summary>Starts a run; false when one is already running.</summary>
    public bool Start(DateTime from)
    {
        CancellationToken token;
        lock (_gate)
        {
            if (IsRunning) return false;
            IsRunning = true;
            StartedAt = DateTime.Now;
            FinishedAt = null;
            Status = "Running";
            _log.Clear();
            _cts = new CancellationTokenSource();
            token = _cts.Token;
        }

        _ = Task.Run(() => RunAsync(from.Date, token));
        Changed?.Invoke();
        return true;
    }

    public void Cancel()
    {
        lock (_gate) _cts?.Cancel();
    }

    private async Task RunAsync(DateTime from, CancellationToken token)
    {
        var progress = new InlineProgress(Append);
        DateTime to = HistoryCachePolicy.LastHistoryDay(DateTime.Today);
        string status;
        try
        {
            Append($"History cache {from:yyyy-MM-dd} – {to:yyyy-MM-dd} (the last {HistoryCachePolicy.LiveWindowDays} days stay live).");
            using IServiceScope scope = scopes.CreateScope();

            var ng = scope.ServiceProvider.GetRequiredService<NgRateService>();
            HistoryCacheResult ngResult = await ng.BuildHistoryCacheAsync(from, progress, token);

            token.ThrowIfCancellationRequested();
            int fcostFailed = 0;
            if (from <= to)
            {
                var fcost = scope.ServiceProvider.GetRequiredService<FCostService>();
                Append("F-COST RAW: fetching one day per request…");
                FCostRawBackfillResult fcostResult = await fcost.BackfillRawAsync(
                    from, to, delayMs: FCostDelayMs, progress: progress, cancellationToken: token);
                fcostFailed = fcostResult.FailedDays;
            }

            int failed = ngResult.FailedDays + fcostFailed;
            status = failed == 0 ? "Completed" : $"Completed with {failed:N0} failed day(s) — run again to retry them";
        }
        catch (OperationCanceledException)
        {
            status = "Stopped — run again to continue from where it stopped";
        }
        catch (Exception ex)
        {
            status = $"Failed: {ex.Message}";
        }

        Append(status);
        lock (_gate)
        {
            Status = status;
            IsRunning = false;
            FinishedAt = DateTime.Now;
            _cts?.Dispose();
            _cts = null;
        }
        Changed?.Invoke();
    }

    private void Append(string line)
    {
        activity.Log("HistoryCache", line);
        lock (_gate)
        {
            _log.Add($"{DateTime.Now:HH:mm:ss}  {line}");
            if (_log.Count > MaxLogLines) _log.RemoveRange(0, _log.Count - MaxLogLines);
        }
        Changed?.Invoke();
    }

    /// <summary>Calls back on the reporting thread; Progress&lt;T&gt; would post to the thread pool
    /// and reorder lines.</summary>
    private sealed class InlineProgress(Action<string> report) : IProgress<string>
    {
        public void Report(string value) => report(value);
    }
}
