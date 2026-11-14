namespace PortalGNLAB.Web.Services;

/// <summary>The Process List (공정편성 LIST) of each model.</summary>
public sealed class WorkingProcessListService(string? root = null)
{
    private string Root => root ?? AppStoragePaths.AppFolder("working-process-list");
}
