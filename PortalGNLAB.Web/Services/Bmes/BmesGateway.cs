namespace PortalGNLAB.Web.Services;

/// <summary>
/// Every BMES data pull in the app, in one class. Feature services (NG Rate, F-Cost,
/// Report, Worker Status, LPA, ...) must not call the BMES server themselves: they open a
/// <see cref="BmesWebSession"/> and ask this gateway for a data set, then store and
/// process what comes back. When BMES provides an official API, only the files in
/// Services/Bmes change; each method keeps returning the same shape the features use.
///
/// Data source per screen: PLATFORM_API Tool through <see cref="BmesApi"/> (api) or the
/// BMES web page through <see cref="BmesWebSession"/> (web; the web login is lazy):
///
/// One partial file per BMES screen (BmesGateway.&lt;Screen&gt;.cs):
///
///   Screen      Data                          Source                      Used by
///   MES020210   NG records (input / NG qty)   api  AI_TOOL_00055 view S   NgRateService
///   MES072400   F-Cost by model / material    api  AI_TOOL_00001 view P   FCostService
///   MES072410   Core-parts F-Cost             api  AI_TOOL_00001 view G   FCostCorePartsService
///   MES050032   IPG defects (week / month)    web  (no Tool)              IpgDefectService
///   MES072900   Worker attendance             web  (AI_TOOL_00005 is "HR Personnel Report", not
///                                                  authorised for this login)   WorkerStatusService
///   MES020010   Material master               web  (no Tool)              BmesMaterialService
///   MES020240   Routing master                web  (AI_TOOL_00028 is MES020030/033: per version,
///                                                  not the same rows)     BmesRoutingScrapeService
///   MES073300   Drawing / spec (PDM)          web  (no Tool)              BmesPdmScrapeService
///   MES073260   LPA audit list                api  AI_TOOL_00027 view C   BmesLpaScrapeService
///   MES073261   LPA checklist                 api  AI_TOOL_00027 view D   BmesLpaScrapeService
///   MES073261   LPA photos (GetImage)         web  (anonymous download)   BmesLpaImageService
///
/// Not here yet: the read-only SQL queries against the BMES database (MATE, MAST/BOMC,
/// INFR, STBI, TCURR, E008) in <see cref="BmesFcostActualService"/>.
/// </summary>
public static partial class BmesGateway
{
    /// <summary>Logs the session in to the BMES web when a web-parsed screen needs it. Null = ready,
    /// otherwise the reason (same wording the services showed when the login used to be eager).</summary>
    private static async Task<string?> WebLoginErrorAsync(
        BmesWebSession session, IProgress<string>? progress, CancellationToken cancellationToken)
    {
        BmesLoginOutcome outcome = await session.EnsureWebLoginAsync(progress, cancellationToken);
        if (outcome.Succeeded) return null;
        return outcome.Failure switch
        {
            BmesLoginFailure.Token => "Failed to read BMES login token.",
            BmesLoginFailure.MissingCredentials => outcome.Diagnostic,
            _ => $"BMES login failed - {outcome.Diagnostic}",
        };
    }

    /// <summary>The API credentials of the session, or the reason there are none.</summary>
    private static BmesApiCredentials RequireApiCredentials(BmesWebSession session) =>
        session.ApiCredentials
        ?? throw new BmesApiException(0, "NO_CREDENTIALS",
            "BMES API login is not configured. Set the API facility / id / password (or the BMES login) in Settings.");
}
