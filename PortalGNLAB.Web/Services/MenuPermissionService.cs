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
        GrantWorkingOnce();
    }

    private const string WorkingGrantMarker = "Migration:grant-working-2026-11-09";

    /// <summary>Roles seeded before the Working menu existed never get it from <c>DefaultsByRole</c>, so grant it once to the default roles that already have permissions. The marker keeps a later admin removal.</summary>
    private void GrantWorkingOnce()
    {
        if (!string.IsNullOrEmpty(_repo.GetSetting(WorkingGrantMarker))) return;
        foreach ((string role, string[] menus) in AppMenus.DefaultsByRole)
        {
            if (role == AppRoles.Admin || !menus.Contains(AppMenus.Working)) continue;
            if (_repo.GetMenuPermissionsForRole(role).Count > 0)
                _repo.SetMenuPermission(role, AppMenus.Working, true);
        }
        _repo.SetSetting(WorkingGrantMarker, DateTime.UtcNow.ToString("O"));
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
