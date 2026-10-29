using System.Text.RegularExpressions;

namespace PortalGNLAB.Web.Services;

/// <summary>One routing step: material × process with its line stage (ZLRTP = routing WPHD.LGUBN).</summary>
public sealed record BmesRoute(string Werks, string Matnr, string ProcessCode, string ProcessName, string StageCode, string StageName);

/// <summary>One defect type: process × NG code with their names.</summary>
public sealed record BmesNgCode(string ProcessCode, string ProcessName, string NgCode, string NgName);

/// <param name="UnnamedRoutes">routing steps left out because AI_TOOL_00056 has no process name</param>
/// <param name="UnnamedNgCodes">process × NG codes left out because AI_TOOL_00056 has no name for them</param>
public sealed record BmesRoutingSource(List<BmesRoute> Routes, List<BmesNgCode> NgCodes, int UnnamedRoutes, int UnnamedNgCodes);

public static partial class BmesGateway
{
    /// <summary>Factory whose raw data is pulled (plants 3200 / 3210 / 3220).</summary>
    public const string RoutingFactory = "GN";

    /// <summary>Line stage IPG = the process is not in the routing (spec AI_TOOL_00055 rule 11).</summary>
    private const string NoRoutingStage = "L99";

    /// <summary>AI_TOOL_00055 view R (raw data, GN): result 1 input per W/O × material × process with ZLRTP,
    /// result 2 NG per W/O × process × NG code. At most 3 months per call.</summary>
    public static Dictionary<string, object> RoutingRawParameters(DateTime from, DateTime to) => new()
    {
        ["ZTYPE"] = "R",
        ["FRDAT"] = from.ToString("yyyy-MM-dd"),
        ["STDAT"] = to.ToString("yyyy-MM-dd"),
        ["FACCO"] = new[] { RoutingFactory },
    };

    /// <summary>AI_TOOL_00056: process names (KTSCH), NG code names (ZCODE; its process name is filled only when
    /// KTSCH is asked in the same call) and line stage names (ZLRTP). SDATE = processes / codes seen since then.</summary>
    public static Dictionary<string, object> CodeMasterParameters(DateTime from) => new()
    {
        ["CODTP"] = new[] { "KTSCH", "ZCODE", "ZLRTP" },
        ["SDATE"] = from.ToString("yyyy-MM-dd"),
    };

    /// <summary>
    /// View R + code master → routing steps (material × process, line stages other than L99 IPG, smallest stage
    /// code when a process has several) and the process × NG codes that occurred, all named from AI_TOOL_00056.
    /// Checked 2026-10-29 against the MES020240 routing dump: KTSCH = VLSCH, ZLRTP = LGUBN, KTSCH name = VLSCH_TX.
    /// </summary>
    public static BmesRoutingSource RoutingSourceFromResults(BmesApiResult raw, BmesApiResult codes)
    {
        var processNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var stageNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var ngNames = new Dictionary<(string, string), (string Process, string Ng)>();
        if (BmesToolRows.FindSet(codes, "CODTP", "CODNO", "CODNO_TX") is { } codeSet)
        {
            foreach (var row in BmesToolRows.ToRowDictionaries(codeSet))
            {
                string codtp = row["CODTP"].Trim().ToUpperInvariant();
                string code = row["CODNO"].Trim();
                string name = CleanName(row["CODNO_TX"]);
                if (code.Length == 0) continue;
                switch (codtp)
                {
                    case "KTSCH" when name.Length > 0: processNames.TryAdd(code, name); break;
                    case "ZLRTP" when name.Length > 0: stageNames.TryAdd(code, name); break;
                    case "ZCODE":
                        string parent = row.GetValueOrDefault("PCDNO", "").Trim();
                        if (parent.Length > 0)
                            ngNames.TryAdd((parent.ToUpperInvariant(), code.ToUpperInvariant()),
                                (CleanName(row.GetValueOrDefault("PCDNO_TX", "")), name));
                        break;
                }
            }
        }

        var routes = new List<BmesRoute>();
        int unnamedRoutes = 0;
        if (BmesToolRows.FindSet(raw, "MATNR", "KTSCH", "ZLRTP", "INQTY") is { } inputSet)
        {
            var steps = BmesToolRows.ToRowDictionaries(inputSet)
                .Select(r => (Werks: r.GetValueOrDefault("WERKS", "").Trim(), Matnr: r["MATNR"].Trim(),
                    Process: r["KTSCH"].Trim(), Stage: r["ZLRTP"].Trim().ToUpperInvariant()))
                .Where(r => r.Matnr.Length > 0 && r.Process.Length > 0 && r.Stage.Length > 0 && r.Stage != NoRoutingStage)
                .GroupBy(r => (Matnr: r.Matnr.ToUpperInvariant(), Process: r.Process.ToUpperInvariant()));
            foreach (var g in steps)
            {
                var first = g.First();
                if (!processNames.TryGetValue(first.Process, out string? processName)) { unnamedRoutes++; continue; }
                string stage = g.Select(r => r.Stage).Order(StringComparer.Ordinal).First();
                routes.Add(new BmesRoute(first.Werks, first.Matnr, first.Process, processName, stage,
                    stageNames.GetValueOrDefault(stage, stage)));
            }
        }

        var ngCodes = new List<BmesNgCode>();
        int unnamedNgCodes = 0;
        if (BmesToolRows.FindSet(raw, "KTSCH", "ZCODE", "NGQTY") is { } ngSet)
        {
            var pairs = BmesToolRows.ToRowDictionaries(ngSet)
                .Select(r => (Process: r["KTSCH"].Trim(), Code: r["ZCODE"].Trim()))
                .Where(r => r.Process.Length > 0 && r.Code.Length > 0)
                .DistinctBy(r => (r.Process.ToUpperInvariant(), r.Code.ToUpperInvariant()));
            foreach (var (process, code) in pairs)
            {
                if (!ngNames.TryGetValue((process.ToUpperInvariant(), code.ToUpperInvariant()), out var names) || names.Ng.Length == 0)
                {
                    unnamedNgCodes++;
                    continue;
                }
                string processName = processNames.GetValueOrDefault(process) ?? names.Process;
                if (processName.Length == 0) { unnamedNgCodes++; continue; }
                ngCodes.Add(new BmesNgCode(process, processName, code, names.Ng));
            }
        }

        return new BmesRoutingSource(
            routes.OrderBy(r => r.Matnr, StringComparer.Ordinal).ThenBy(r => r.ProcessCode, StringComparer.Ordinal).ToList(),
            ngCodes.OrderBy(r => r.ProcessCode, StringComparer.Ordinal).ThenBy(r => r.NgCode, StringComparer.Ordinal).ToList(),
            unnamedRoutes, unnamedNgCodes);
    }

    /// <summary>
    /// Routing steps and occurred NG codes of the last 3 months up to <paramref name="to"/>, from PLATFORM_API:
    /// AI_TOOL_00055 view R (GN) then AI_TOOL_00056 (names), one call each.
    /// Throws <see cref="BmesApiException"/> when a call fails.
    /// </summary>
    public static async Task<BmesRoutingSource> FetchRoutingSourceAsync(
        BmesWebSession session, DateTime to, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        BmesApiCredentials credentials = RequireApiCredentials(session);
        DateTime from = to.Date.AddMonths(-BmesApi.MaxPeriodMonths).AddDays(1);

        progress?.Report($"Calling AI_TOOL_00055 view R ({RoutingFactory}, {from:yyyy-MM-dd} ~ {to:yyyy-MM-dd})…");
        BmesApiResult raw = await BmesApi.ExecuteAsync(credentials, "AI_TOOL_00055",
            RoutingRawParameters(from, to.Date), cancellationToken: cancellationToken);

        progress?.Report("Calling AI_TOOL_00056 (process / NG code names)…");
        BmesApiResult codes = await BmesApi.ExecuteAsync(credentials, "AI_TOOL_00056",
            CodeMasterParameters(from), cancellationToken: cancellationToken);

        return RoutingSourceFromResults(raw, codes);
    }

    /// <summary>
    /// Material names (MAKTX) for materials of one plant, from AI_TOOL_00055 view S filtered by MATNR
    /// (view R has no material name column). Five materials per call.
    /// </summary>
    public static async Task<Dictionary<string, string>> FetchMaterialNamesAsync(
        BmesWebSession session, string werks, IReadOnlyCollection<string> matnrs, DateTime to,
        CancellationToken cancellationToken = default)
    {
        BmesApiCredentials credentials = RequireApiCredentials(session);
        DateTime from = to.Date.AddMonths(-BmesApi.MaxPeriodMonths).AddDays(1);
        var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string[] chunk in matnrs.Chunk(5))
        {
            var parameters = NgRecordParameters(werks, from, to.Date);
            parameters["MATNR"] = chunk;
            BmesApiResult result = await BmesApi.ExecuteAsync(credentials, "AI_TOOL_00055", parameters,
                cancellationToken: cancellationToken);
            foreach (var row in NgRecordRowsFromResult(result))
            {
                string code = row.GetValueOrDefault("MATERIALCODE", "").Trim();
                string name = row.GetValueOrDefault("MATERIALNAME", "").Trim();
                if (code.Length > 0 && name.Length > 0) names.TryAdd(code, name);
            }
        }
        return names;
    }

    /// <summary>Names can hold line breaks and trailing blanks (spec AI_TOOL_00056) — one space instead.</summary>
    private static string CleanName(string text) => Blanks().Replace(text, " ").Trim();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Blanks();
}
