using System.Security.Claims;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>Working API: model groups (read-only), per-model info and dev / issue / review entries.</summary>
public static class WorkingEndpoints
{
    private static string User(ClaimsPrincipal user) => user.Identity?.Name ?? "";

    public static IEndpointRouteBuilder MapWorking(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/portal/tools/working").RequireAuthorization();
        api.AddEndpointFilter(async (ctx, next) =>
        {
            if (!ctx.HttpContext.RequestServices.GetRequiredService<MenuPermissionService>()
                .IsAllowed(ctx.HttpContext.User, AppMenus.Working)) return Results.Forbid();
            if (!HttpMethods.IsGet(ctx.HttpContext.Request.Method)
                && ctx.HttpContext.Request.Headers["X-Requested-With"] != "PortalGNLAB")
                return Results.BadRequest(new { error = "Required request header is missing." });
            try { return await next(ctx); }
            catch (WorkingConflictException ex) { return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status409Conflict); }
            catch (ArgumentException ex) { return Results.BadRequest(new { error = ex.Message }); }
        });

        api.MapGet("/groups", (WorkingService svc) => Results.Ok(svc.ListGroups()));
        api.MapGet("/models/{name}", (string name, WorkingService svc) => Results.Ok(svc.GetModel(name)));
        api.MapPut("/models/{name}", (string name, WorkingModelSave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.SaveModel(name, body, User(user))));

        api.MapGet("/models/{name}/entries", (string name, WorkingService svc) => Results.Ok(new { entries = svc.ListEntries(name) }));
        api.MapPost("/models/{name}/entries", (string name, WorkingEntrySave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.CreateEntry(name, body, User(user))));
        api.MapPut("/models/{name}/entries/{id}", (string name, string id, WorkingEntrySave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.UpdateEntry(name, id, body, User(user)) is { } e ? Results.Ok(e) : Results.NotFound());
        api.MapDelete("/models/{name}/entries/{id}", (string name, string id, WorkingService svc) =>
            svc.DeleteEntry(name, id) ? Results.NoContent() : Results.NotFound());

        return app;
    }
}
