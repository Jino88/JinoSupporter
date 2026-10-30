using System.Security.Claims;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// Runtime arbiter of which menu items a user (by their role) is allowed to see.
/// Admin role ALWAYS sees everything, regardless of DB state.
/// </summary>
public sealed class MenuPermissionService
{
    private readonly WebRepository _repo;

    public MenuPermissionService(WebRepository repo)
    {
        _repo = repo;
        _repo.SeedDefaultMenuPermissionsIfEmpty();
        GrantProductionInputOnce();
    }

    private const string ProductionInputGrantMarker = "Migration:grant-production-input-2026-10-30";

    /// <summary>
    /// Roles seeded before Production Input existed never get it from <c>DefaultsByRole</c>, so grant it
    /// once to the default roles that already have permissions. The marker keeps a later admin removal.
    /// </summary>
    private void GrantProductionInputOnce()
    {
        if (!string.IsNullOrEmpty(_repo.GetSetting(ProductionInputGrantMarker))) return;
        foreach ((string role, string[] menus) in AppMenus.DefaultsByRole)
        {
            if (role == AppRoles.Admin || !menus.Contains(AppMenus.ProductionInput)) continue;
            if (_repo.GetMenuPermissionsForRole(role).Count > 0)
                _repo.SetMenuPermission(role, AppMenus.ProductionInput, true);
        }
        _repo.SetSetting(ProductionInputGrantMarker, DateTime.UtcNow.ToString("O"));
    }

    public bool IsAllowed(string role, string menuId)
    {
        if (string.Equals(role, AppRoles.Admin, StringComparison.OrdinalIgnoreCase))
            return true;
        HashSet<string> set = _repo.GetMenuPermissionsForRole(role);
        return set.Contains(menuId);
    }

    public bool IsAllowed(ClaimsPrincipal user, string menuId)
    {
        string? role = user.FindFirst(ClaimTypes.Role)?.Value;
        if (string.IsNullOrEmpty(role)) return false;
        return IsAllowed(role, menuId);
    }
}
