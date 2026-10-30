using System.Security.Claims;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>Production Input API: models with their setup, daily entries, plan and Excel export.</summary>
public static class ProductionInputEndpoints
{
    public const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private static string User(ClaimsPrincipal user) => user.Identity?.Name ?? "";

    public static IEndpointRouteBuilder MapProductionInput(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/portal/tools/production-input").RequireAuthorization();
        api.AddEndpointFilter(async (ctx, next) =>
        {
            if (!ctx.HttpContext.RequestServices.GetRequiredService<MenuPermissionService>()
                .IsAllowed(ctx.HttpContext.User, AppMenus.ProductionInput)) return Results.Forbid();
            if (!HttpMethods.IsGet(ctx.HttpContext.Request.Method)
                && ctx.HttpContext.Request.Headers["X-Requested-With"] != "PortalGNLAB")
                return Results.BadRequest(new { error = "Required request header is missing." });
            try { return await next(ctx); }
            catch (ProductionConflictException ex) { return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status409Conflict); }
            catch (ArgumentException ex) { return Results.BadRequest(new { error = ex.Message }); }
        });

        api.MapGet("", (ProductionInputService svc) => Results.Ok(new { models = svc.ListModels() }));
        api.MapGet("/models/{id}", (string id, ProductionInputService svc) =>
            svc.GetModel(id) is { } m ? Results.Ok(m) : Results.NotFound());
        api.MapPost("/models", (ProductionModelSave body, ProductionInputService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.CreateModel(body, User(user))));
        api.MapPut("/models/{id}", (string id, ProductionModelSave body, ProductionInputService svc, ClaimsPrincipal user) =>
            svc.UpdateModel(id, body, User(user)) is { } m ? Results.Ok(m) : Results.NotFound());
        api.MapDelete("/models/{id}", (string id, ProductionInputService svc) =>
            svc.DeleteModel(id) ? Results.NoContent() : Results.NotFound());

        api.MapGet("/models/{id}/days/{date}", (string id, string date, ProductionInputService svc) => Results.Ok(svc.GetDay(id, date)));
        api.MapPut("/models/{id}/days/{date}", (string id, string date, ProductionDaySave body, ProductionInputService svc, ClaimsPrincipal user) =>
            svc.SaveDay(id, date, body, User(user)) is { } d ? Results.Ok(d) : Results.NotFound());
        api.MapDelete("/models/{id}/days/{date}", (string id, string date, ProductionInputService svc) =>
        {
            svc.DeleteDay(id, date);
            return Results.NoContent();
        });
        api.MapGet("/models/{id}/days", (string id, string? from, string? to, ProductionInputService svc) =>
            Results.Ok(svc.ListDays(id, from, to)));

        api.MapGet("/models/{id}/plan", (string id, ProductionInputService svc) => Results.Ok(svc.GetPlan(id)));
        api.MapPut("/models/{id}/plan", (string id, ProductionPlanSave body, ProductionInputService svc, ClaimsPrincipal user) =>
            svc.SavePlan(id, body, User(user)) is { } p ? Results.Ok(p) : Results.NotFound());

        api.MapPost("/export", (ExportRequest body) =>
            Results.File(ProductionInputService.ExportXlsx(body), XlsxContentType, ProductionInputService.SafeFileName(body.FileName)));

        return app;
    }
}
