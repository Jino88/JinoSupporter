namespace PortalGNLAB.Web.Services;

/// <summary>The Work Instructions (작업지도서) of each model.</summary>
public sealed class WorkingInstructionService(string? root = null)
{
    private string Root => root ?? AppStoragePaths.AppFolder("working-instructions");
}
