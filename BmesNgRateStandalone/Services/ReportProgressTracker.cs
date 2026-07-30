using System.Diagnostics;

namespace BmesNgRateStandalone.Services;

public enum ReportStageStatus { Pending, Running, Done, Failed }

/// <summary>
/// Optional capability on a progress sink: report "n of m" alongside the text.
///
/// Added as a capability rather than a new <c>IProgress&lt;T&gt;</c> signature so the
/// existing <c>IProgress&lt;string&gt;</c> plumbing — shared by every NG-rate report
/// page — keeps working untouched. Sinks that do not implement this simply ignore the
/// counts.
/// </summary>
public interface IStepProgress
{
    void ReportSteps(int current, int total, string? phase = null);
}

public static class StepProgressExtensions
{
    /// <summary>No-op when the sink cannot use counts, so producers can always call it.</summary>
    public static void ReportSteps(
        this IProgress<string>? progress, int current, int total, string? phase = null)
    {
        if (progress is IStepProgress sink)
            sink.ReportSteps(current, total, phase);
    }
}

/// <summary>One tracked stage of a multi-part report run (Daily, Weekly, …).</summary>
public sealed class ReportStage
{
    public required string Name { get; init; }

    public ReportStageStatus Status { get; internal set; } = ReportStageStatus.Pending;

    /// <summary>Latest progress message, shown as the stage's caption.</summary>
    public string CurrentStep { get; internal set; } = string.Empty;

    /// <summary>Progress messages seen so far in this run.</summary>
    public int StepCount { get; internal set; }

    /// <summary>How many messages this stage emitted on its last successful run — the
    /// denominator for <see cref="Percent"/>. Null until it has completed once.</summary>
    public int? BaselineSteps { get; internal set; }

    public TimeSpan Elapsed { get; internal set; }

    /// <summary>Counted work reported by the producer, e.g. 5 of 9 fetch ranges. Present
    /// only while a counted phase is running; this is a real measurement, not an
    /// estimate, so it takes priority over <see cref="BaselineSteps"/>.</summary>
    public int? StepCurrent { get; internal set; }
    public int? StepTotal   { get; internal set; }
    public string? PhaseName { get; internal set; }

    private double _percent;

    /// <summary>Null while running with nothing to measure against — neither counted
    /// steps nor a baseline — because there is no honest number to show; the bar
    /// animates instead.
    ///
    /// Held below 100 until the stage reports Done, and never allowed to move backwards
    /// within a run: a bar that retreats reads as a bug.</summary>
    public double? Percent => Status switch
    {
        ReportStageStatus.Done                 => 100d,
        ReportStageStatus.Running when _percent > 0 => _percent,
        _                                      => null,
    };

    internal void Advance(double? candidate)
    {
        if (candidate is not double value) return;
        _percent = Math.Max(_percent, Math.Clamp(value, 2d, 95d));
    }

    internal void ResetRunState()
    {
        Status      = ReportStageStatus.Pending;
        CurrentStep = string.Empty;
        StepCount   = 0;
        StepCurrent = null;
        StepTotal   = null;
        PhaseName   = null;
        Elapsed     = TimeSpan.Zero;
        _percent    = 0;
    }

    /// <summary>Estimate from the last successful run; used only when the producer is
    /// not reporting counted steps.</summary>
    internal double? BaselineEstimate =>
        BaselineSteps is > 0 ? StepCount * 100d / BaselineSteps.Value : null;
}

/// <summary>
/// Turns the free-text <see cref="IProgress{T}"/> streams of a multi-stage report run
/// into per-stage status, so the UI can show a few progress bars instead of a wall of
/// log lines.
///
/// Progress callbacks arrive from background tasks running in parallel, so every
/// mutation is locked.
/// </summary>
public sealed class ReportProgressTracker
{
    private readonly object _gate = new();
    private readonly List<ReportStage> _stages;
    private readonly Dictionary<string, Stopwatch> _watches = new(StringComparer.Ordinal);
    private readonly Action? _onChanged;

    public ReportProgressTracker(IEnumerable<string> stageNames, Action? onChanged = null)
    {
        _stages = stageNames.Select(n => new ReportStage { Name = n }).ToList();
        _onChanged = onChanged;
    }

    public IReadOnlyList<ReportStage> Stages => _stages;

    /// <summary>Whole-run completion: finished stages count fully, the running one
    /// contributes its own fraction.</summary>
    public double OverallPercent
    {
        get
        {
            lock (_gate)
            {
                if (_stages.Count == 0) return 0;
                double sum = _stages.Sum(s => s.Status switch
                {
                    ReportStageStatus.Done   => 100d,
                    ReportStageStatus.Failed => 100d,
                    _                        => s.Percent ?? 0d,
                });
                return sum / _stages.Count;
            }
        }
    }

    public bool IsComplete
    {
        get
        {
            lock (_gate)
                return _stages.All(s => s.Status is ReportStageStatus.Done or ReportStageStatus.Failed);
        }
    }

    /// <summary>Clear run state, keeping the learned baselines.</summary>
    public void Reset()
    {
        lock (_gate)
        {
            foreach (ReportStage stage in _stages)
                stage.ResetRunState();
            _watches.Clear();
        }
        _onChanged?.Invoke();
    }

    /// <summary>An <see cref="IProgress{T}"/> that advances <paramref name="stageName"/>
    /// and forwards every message to <paramref name="alsoLog"/> for the detail log.</summary>
    public IProgress<string> For(string stageName, Action<string>? alsoLog = null) =>
        new StageProgress(this, stageName, alsoLog);

    public void MarkRunning(string stageName) => Transition(stageName, ReportStageStatus.Running);
    public void MarkDone(string stageName)    => Transition(stageName, ReportStageStatus.Done);
    public void MarkFailed(string stageName)  => Transition(stageName, ReportStageStatus.Failed);

    /// <summary>Any stage still pending or running when the run ends never reported a
    /// result — settle it so the bars do not sit half-drawn forever.</summary>
    public void SettleUnfinished(bool failed)
    {
        lock (_gate)
        {
            foreach (ReportStage stage in _stages)
            {
                if (stage.Status is ReportStageStatus.Done or ReportStageStatus.Failed) continue;
                stage.Status = failed ? ReportStageStatus.Failed : ReportStageStatus.Done;
                StopWatch(stage);
            }
        }
        _onChanged?.Invoke();
    }

    private void Transition(string stageName, ReportStageStatus status)
    {
        lock (_gate)
        {
            ReportStage? stage = Find(stageName);
            if (stage is null) return;

            if (status == ReportStageStatus.Running)
            {
                if (stage.Status == ReportStageStatus.Running) return;
                stage.Status = ReportStageStatus.Running;
                _watches[stage.Name] = Stopwatch.StartNew();
            }
            else
            {
                stage.Status = status;
                // A completed run calibrates the next one's percentage.
                if (status == ReportStageStatus.Done && stage.StepCount > 0)
                    stage.BaselineSteps = stage.StepCount;
                StopWatch(stage);
            }
        }
        _onChanged?.Invoke();
    }

    private void StopWatch(ReportStage stage)
    {
        if (!_watches.TryGetValue(stage.Name, out Stopwatch? sw)) return;
        sw.Stop();
        stage.Elapsed = sw.Elapsed;
    }

    private ReportStage? Find(string name) =>
        _stages.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.Ordinal));

    private void Advance(string stageName, string message)
    {
        lock (_gate)
        {
            ReportStage? stage = Find(stageName);
            if (stage is null) return;

            if (stage.Status == ReportStageStatus.Pending)
            {
                stage.Status = ReportStageStatus.Running;
                _watches[stage.Name] = Stopwatch.StartNew();
            }

            stage.StepCount++;
            stage.CurrentStep = Shorten(message);

            // Counted steps win; the baseline estimate only fills the gaps between
            // (and after) counted phases.
            if (stage.StepTotal is null)
                stage.Advance(stage.BaselineEstimate);

            if (_watches.TryGetValue(stage.Name, out Stopwatch? sw))
                stage.Elapsed = sw.Elapsed;
        }
        _onChanged?.Invoke();
    }

    private void AdvanceSteps(string stageName, int current, int total, string? phase)
    {
        lock (_gate)
        {
            ReportStage? stage = Find(stageName);
            if (stage is null || total <= 0) return;

            if (stage.Status == ReportStageStatus.Pending)
            {
                stage.Status = ReportStageStatus.Running;
                _watches[stage.Name] = Stopwatch.StartNew();
            }

            stage.PhaseName = phase;
            stage.Advance(current * 100d / total);

            if (current >= total)
            {
                // Phase finished — stop showing "9/9" and let later messages drive the
                // bar again, from the level this phase reached.
                stage.StepCurrent = null;
                stage.StepTotal   = null;
                stage.PhaseName   = null;
            }
            else
            {
                stage.StepCurrent = current;
                stage.StepTotal   = total;
            }

            if (_watches.TryGetValue(stage.Name, out Stopwatch? sw))
                stage.Elapsed = sw.Elapsed;
        }
        _onChanged?.Invoke();
    }

    /// <summary>Progress lines carry bracketed sub-labels and trailing timings that make
    /// a one-line caption unreadable; keep the human part.</summary>
    private static string Shorten(string message)
    {
        string text = message.Trim();
        while (text.StartsWith('[') && text.IndexOf(']') > 0)
            text = text[(text.IndexOf(']') + 1)..].TrimStart();

        text = text.TrimStart('─', '-', ' ');
        return text.Length <= 90 ? text : text[..90] + "…";
    }

    /// <summary>Baselines survive across page visits so the second run onward shows a
    /// real percentage. Caller owns the storage.</summary>
    public Dictionary<string, int> ExportBaselines()
    {
        lock (_gate)
            return _stages
                .Where(s => s.BaselineSteps is > 0)
                .ToDictionary(s => s.Name, s => s.BaselineSteps!.Value, StringComparer.Ordinal);
    }

    public void ImportBaselines(IReadOnlyDictionary<string, int>? baselines)
    {
        if (baselines is null) return;
        lock (_gate)
        {
            foreach (ReportStage stage in _stages)
                if (baselines.TryGetValue(stage.Name, out int steps) && steps > 0)
                    stage.BaselineSteps = steps;
        }
    }

    private sealed class StageProgress(ReportProgressTracker owner, string stageName, Action<string>? alsoLog)
        : IProgress<string>, IStepProgress
    {
        public void Report(string value)
        {
            alsoLog?.Invoke(value);
            owner.Advance(stageName, value);
        }

        public void ReportSteps(int current, int total, string? phase = null)
            => owner.AdvanceSteps(stageName, current, total, phase);
    }
}
