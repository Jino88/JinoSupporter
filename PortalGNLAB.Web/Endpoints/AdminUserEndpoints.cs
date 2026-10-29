using System.Security.Claims;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>
/// User management for the React page /app/admin/users (port of Components/Pages/UsersPage.razor).
/// Admin only. Changes require the X-Requested-With header so a cross-site form cannot post them.
/// </summary>
public static class AdminUserEndpoints
{
    public const string RequestedWithHeader = "X-Requested-With";
    public const string RequestedWithValue = "PortalGNLAB";

    public sealed record AdminUser(long Id, string Username, string DisplayName, string Role, string CreatedAt);

    public sealed record AdminMenu(string Id, string Label, string Group);

    public sealed record AdminUsersView(
        string CurrentUser,
        IReadOnlyList<string> Roles,
        IReadOnlyList<AdminUser> Users,
        IReadOnlyList<AdminMenu> Menus,
        IReadOnlyDictionary<string, string[]> Permissions);

    public sealed record NewUser(string Username, string DisplayName, string Password, string Role);

    public sealed record RoleChange(string Role);

    public sealed record NameChange(string DisplayName);

    public sealed record PasswordChange(string Password);

    public sealed record PermissionChange(string Role, string MenuId, bool Allowed);

    /// <summary>Checks a new user. Returns the error text to show, or null when it can be added.</summary>
    public static string? ValidateNewUser(NewUser user)
    {
        if (string.IsNullOrWhiteSpace(user.Username)) return "Username required.";
        if (string.IsNullOrWhiteSpace(user.Password)) return "Password required.";
        if (!AppRoles.All.Contains(user.Role)) return "Unknown role.";
        return null;
    }

    private static bool IsSelf(ClaimsPrincipal user, UserRecord target) =>
        target.Username.Equals(user.Identity?.Name ?? "", StringComparison.OrdinalIgnoreCase);

    public static IEndpointRouteBuilder MapAdminUsers(this IEndpointRouteBuilder app)
    {
        var admin = app.MapGroup("/api/portal/admin").RequireAuthorization(p => p.RequireRole(AppRoles.Admin));

        // Mutations must come from the React app (custom header → browsers preflight cross-site requests).
        admin.AddEndpointFilter(async (ctx, next) =>
        {
            var request = ctx.HttpContext.Request;
            if (!HttpMethods.IsGet(request.Method) &&
                request.Headers[RequestedWithHeader] != RequestedWithValue)
                return Results.StatusCode(StatusCodes.Status400BadRequest);
            return await next(ctx);
        });

        admin.MapGet("/users", (ClaimsPrincipal user, WebRepository repo) => new AdminUsersView(
            user.Identity?.Name ?? "",
            AppRoles.All,
            repo.GetAllUsers().Select(u => new AdminUser(u.Id, u.Username, u.DisplayName, u.Role, u.CreatedAt)).ToList(),
            AppMenus.All.Select(m => new AdminMenu(m.Id, m.Label, m.Group)).ToList(),
            repo.GetAllMenuPermissions().ToDictionary(kv => kv.Key, kv => kv.Value.Order().ToArray())));

        admin.MapPost("/users", (NewUser body, WebRepository repo) =>
        {
            if (ValidateNewUser(body) is { } error) return Results.BadRequest(new { error });
            if (repo.GetUser(body.Username.Trim()) is not null) return Results.BadRequest(new { error = "Username already exists." });
            repo.AddUser(body.Username.Trim(), AuthService.HashPassword(body.Password), body.Role, body.DisplayName.Trim());
            return Results.NoContent();
        });

        admin.MapPut("/users/{id:long}/role", (long id, RoleChange body, ClaimsPrincipal user, WebRepository repo) =>
        {
            var target = repo.GetAllUsers().FirstOrDefault(u => u.Id == id);
            if (target is null) return Results.NotFound();
            if (IsSelf(user, target)) return Results.BadRequest(new { error = "You cannot change your own role." });
            if (!AppRoles.All.Contains(body.Role)) return Results.BadRequest(new { error = "Unknown role." });
            repo.UpdateUserRole(id, body.Role);
            return Results.NoContent();
        });

        admin.MapPut("/users/{id:long}/name", (long id, NameChange body, WebRepository repo, ConnectedUsersService connected) =>
        {
            var target = repo.GetAllUsers().FirstOrDefault(u => u.Id == id);
            if (target is null) return Results.NotFound();
            string name = body.DisplayName.Trim();
            repo.UpdateUserDisplayName(id, name);
            connected.UpdateNameByUsername(target.Username, string.IsNullOrEmpty(name) ? target.Username : name);
            return Results.NoContent();
        });

        admin.MapPut("/users/{id:long}/password", (long id, PasswordChange body, WebRepository repo) =>
        {
            if (string.IsNullOrWhiteSpace(body.Password)) return Results.BadRequest(new { error = "Password required." });
            if (repo.GetAllUsers().All(u => u.Id != id)) return Results.NotFound();
            repo.UpdateUserPassword(id, AuthService.HashPassword(body.Password));
            return Results.NoContent();
        });

        admin.MapDelete("/users/{id:long}", (long id, ClaimsPrincipal user, WebRepository repo) =>
        {
            var target = repo.GetAllUsers().FirstOrDefault(u => u.Id == id);
            if (target is null) return Results.NotFound();
            if (IsSelf(user, target)) return Results.BadRequest(new { error = "You cannot remove your own account." });
            repo.DeleteUser(id);
            return Results.NoContent();
        });

        admin.MapPut("/permissions", (PermissionChange body, WebRepository repo) =>
        {
            if (body.Role == AppRoles.Admin) return Results.BadRequest(new { error = "The Admin role always has every menu." });
            if (!AppRoles.All.Contains(body.Role) || AppMenus.All.All(m => m.Id != body.MenuId))
                return Results.BadRequest(new { error = "Unknown role or menu." });
            repo.SetMenuPermission(body.Role, body.MenuId, body.Allowed);
            return Results.NoContent();
        });

        return app;
    }
}
