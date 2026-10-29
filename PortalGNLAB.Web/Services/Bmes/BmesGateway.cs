namespace PortalGNLAB.Web.Services;

/// <summary>
/// Every BMES data pull in the app, in one class. Feature services (NG Rate, F-Cost,
/// Report, Worker Status, LPA, ...) must not call the BMES server themselves: they open a
/// <see cref="BmesWebSession"/> and ask this gateway for a data set, then store and
/// process what comes back. When BMES provides an official API, only the files in
/// Services/Bmes change; each method keeps returning the same shape the features use.
///
/// One partial file per BMES screen (BmesGateway.&lt;Screen&gt;.cs):
///
///   Screen      Data                          Used by
///   MES020210   NG records (input / NG qty)   NgRateService
///   MES072400   F-Cost by model / material    FCostService
///   MES072410   Core-parts F-Cost             FCostCorePartsService
///   MES050032   IPG defects (week / month)    IpgDefectService
///   MES072900   Worker attendance             WorkerStatusService
///   MES020010   Material master               BmesMaterialService
///   MES020240   Routing master                BmesRoutingScrapeService
///   MES073300   Drawing / spec (PDM)          BmesPdmScrapeService
///   MES073260   LPA audit list                BmesLpaScrapeService
///   MES073261   LPA checklist and photos      BmesLpaScrapeService, BmesLpaImageService
///
/// Not here yet: the read-only SQL queries against the BMES database (MATE, MAST/BOMC,
/// INFR, STBI, TCURR, E008) in <see cref="BmesFcostActualService"/>.
/// </summary>
public static partial class BmesGateway
{
}
