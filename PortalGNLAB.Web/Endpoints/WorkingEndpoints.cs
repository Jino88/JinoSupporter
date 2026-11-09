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

        // pictures pasted into an entry (no logId) or into one of its notes
        api.MapPost("/models/{name}/entries/{id}/images", async (string name, string id, HttpContext ctx, WorkingService svc, ClaimsPrincipal user) =>
        {
            (string? logId, List<(string Name, byte[] Data)> files) = await ReadPictures(ctx);
            return svc.AddImages(name, id, logId, files, User(user)) is { } images ? Results.Ok(new { images }) : Results.NotFound();
        });
        api.MapGet("/images/{imageId}", (string imageId, WorkingService svc, HttpContext ctx) =>
        {
            if (svc.GetImage(imageId) is not { } file) return Results.NotFound();
            ctx.Response.Headers.XContentTypeOptions = "nosniff";
            return Results.File(file.Data, file.ContentType);
        });
        api.MapDelete("/models/{name}/entries/{id}/images/{imageId}", (string name, string id, string imageId, WorkingService svc) =>
            svc.DeleteImage(name, id, imageId) ? Results.NoContent() : Results.NotFound());

        return app;
    }

    /// <summary>The "files" and optional "logId" of a multipart upload; the form is read here (no IFormFile binding).</summary>
    private static async Task<(string? LogId, List<(string Name, byte[] Data)> Files)> ReadPictures(HttpContext ctx)
    {
        if (!ctx.Request.HasFormContentType) throw new ArgumentException("Send the pictures as multipart/form-data.");
        const long totalLimit = WorkingService.MaxImagesPerUpload * WorkingService.MaxImageBytes + 1024 * 1024;
        var limit = ctx.Features.Get<IHttpMaxRequestBodySizeFeature>();
        if (limit is { IsReadOnly: false }) limit.MaxRequestBodySize = totalLimit;
        ctx.Features.Set<IFormFeature>(new FormFeature(ctx.Request, new FormOptions { MultipartBodyLengthLimit = totalLimit }));
        IFormCollection form;
        try { form = await ctx.Request.ReadFormAsync(ctx.RequestAborted); }
        catch (InvalidDataException) { throw new ArgumentException("The pictures are too large."); }
        if (form.Files.Count > WorkingService.MaxImagesPerUpload)
            throw new ArgumentException($"Add at most {WorkingService.MaxImagesPerUpload} pictures at a time.");
        var files = new List<(string, byte[])>();
        foreach (IFormFile file in form.Files)
        {
            if (file.Length > WorkingService.MaxImageBytes) throw new ArgumentException($"{file.FileName} is larger than 10 MB.");
            using var data = new MemoryStream();
            await file.CopyToAsync(data, ctx.RequestAborted);
            files.Add((file.FileName, data.ToArray()));
        }
        return (form["logId"].ToString(), files);
    }
}
