using System.Security.Claims;
using Microsoft.AspNetCore.Http.Features;
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

        // dated progress notes under an entry (issues)
        api.MapPost("/models/{name}/entries/{id}/logs", (string name, string id, WorkingLogSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.CreateLog(name, id, body, User(user)) is { } l ? Results.Ok(l) : Results.NotFound());
        api.MapPut("/models/{name}/entries/{id}/logs/{logId}", (string name, string id, string logId, WorkingLogSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.UpdateLog(name, id, logId, body, User(user)) is { } l ? Results.Ok(l) : Results.NotFound());
        api.MapDelete("/models/{name}/entries/{id}/logs/{logId}", (string name, string id, string logId, WorkingService svc) =>
            svc.DeleteLog(name, id, logId) ? Results.NoContent() : Results.NotFound());

        // Setup Plan rows; their problems are issues with a SetupRowId
        api.MapGet("/models/{name}/setup", (string name, WorkingService svc) => Results.Ok(new { rows = svc.ListSetupRows(name) }));
        api.MapPost("/models/{name}/setup", (string name, WorkingSetupRowSave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.CreateSetupRow(name, body, User(user))));
        api.MapPost("/models/{name}/setup/import", (string name, WorkingSetupImport body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(new { rows = svc.ImportSetupRows(name, body.Rows ?? [], User(user)) }));
        api.MapPut("/models/{name}/setup/{rowId}", (string name, string rowId, WorkingSetupRowSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.UpdateSetupRow(name, rowId, body, User(user)) is { } r ? Results.Ok(r) : Results.NotFound());
        api.MapDelete("/models/{name}/setup/{rowId}", (string name, string rowId, WorkingService svc) =>
            svc.DeleteSetupRow(name, rowId) ? Results.NoContent() : Results.NotFound());

        // pictures and videos pasted into an entry (no logId) or into one of its notes, stored as sent
        api.MapPost("/models/{name}/entries/{id}/images", async (string name, string id, HttpContext ctx, WorkingService svc, ClaimsPrincipal user) =>
        {
            (string? logId, List<WorkingUpload> files) = await ReadPictures(ctx);
            return svc.AddFiles(name, id, logId, files, User(user)) is { } images ? Results.Ok(new { images }) : Results.NotFound();
        });
        // streamed from the DB with byte ranges, so a video can be played from any point
        api.MapGet("/images/{imageId}", (string imageId, WorkingService svc, HttpContext ctx) =>
        {
            if (svc.OpenImage(imageId) is not { } file) return Results.NotFound();
            ctx.Response.RegisterForDispose(file);
            ctx.Response.Headers.XContentTypeOptions = "nosniff";
            return Results.Stream(file.Data, file.ContentType, enableRangeProcessing: true);
        });
        api.MapDelete("/models/{name}/entries/{id}/images/{imageId}", (string name, string id, string imageId, WorkingService svc) =>
            svc.DeleteImage(name, id, imageId) ? Results.NoContent() : Results.NotFound());

        return app;
    }

    /// <summary>
    /// The "files" and optional "logId" of a multipart upload; the form is read here (no IFormFile binding). Files over the
    /// in-memory threshold are buffered to a temp file by ASP.NET Core, and each is read from there when stored.
    /// </summary>
    private static async Task<(string? LogId, List<WorkingUpload> Files)> ReadPictures(HttpContext ctx)
    {
        if (!ctx.Request.HasFormContentType) throw new ArgumentException("Send the files as multipart/form-data.");
        const long totalLimit = WorkingService.MaxUploadBytes + 1024 * 1024;
        var limit = ctx.Features.Get<IHttpMaxRequestBodySizeFeature>();
        if (limit is { IsReadOnly: false }) limit.MaxRequestBodySize = totalLimit;
        ctx.Features.Set<IFormFeature>(new FormFeature(ctx.Request, new FormOptions { MultipartBodyLengthLimit = totalLimit }));
        IFormCollection form;
        try { form = await ctx.Request.ReadFormAsync(ctx.RequestAborted); }
        catch (InvalidDataException) { throw new ArgumentException("The files are too large (1 GB per upload)."); }
        if (form.Files.Count > WorkingService.MaxImagesPerUpload)
            throw new ArgumentException($"Add at most {WorkingService.MaxImagesPerUpload} files at a time.");
        var files = form.Files.Select(f => new WorkingUpload(f.FileName, f.Length, f.OpenReadStream)).ToList();
        return (form["logId"].ToString(), files);
    }
}
