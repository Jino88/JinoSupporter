namespace JinoSupporter.Web.Services;

public static class AppRoles
{
    public const string Admin   = "Admin";
    public const string Manager = "Manager";
    public const string Leader  = "Leader";
    public const string Editor  = "Editor";
    public const string Viewer  = "Viewer";

    public static readonly string[] All = [Admin, Manager, Leader, Editor, Viewer];

    // Can input data & use AI extract/tags
    public static bool CanEdit(System.Security.Claims.ClaimsPrincipal user)
        => user.IsInRole(Admin) || user.IsInRole(Manager) || user.IsInRole(Leader) || user.IsInRole(Editor);

    // Can generate AI report
    public static bool CanAiReport(System.Security.Claims.ClaimsPrincipal user)
        => user.IsInRole(Admin) || user.IsInRole(Manager) || user.IsInRole(Leader);

    // Can manage users
    public static bool CanManageUsers(System.Security.Claims.ClaimsPrincipal user)
        => user.IsInRole(Admin);
}
