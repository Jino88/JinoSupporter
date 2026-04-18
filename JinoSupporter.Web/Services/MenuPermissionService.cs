using System.Security.Claims;

namespace JinoSupporter.Web.Services;

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
