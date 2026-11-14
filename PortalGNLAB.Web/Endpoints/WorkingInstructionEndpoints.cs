using System.Security.Claims;
using Microsoft.AspNetCore.Http.Features;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>Working > Work Instructions (작업지도서): mapped inside the Working API group (auth and header checks there).</summary>
public static class WorkingInstructionEndpoints
{
    private const string XlsxType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public static RouteGroupBuilder MapWorkingInstructions(this RouteGroupBuilder api)
    {
        api.MapGet("/models/{name}/instructions", (string name, WorkingInstructionService svc) =>
            svc.Get(WorkingEndpoints.ModelName(name)) is { } doc ? Results.Ok(doc) : Results.NotFound());
        api.MapPost("/models/{name}/instructions", async (string name, HttpContext ctx, WorkingInstructionService svc, ClaimsPrincipal user) =>
        {
            IFormFile file = await ReadWorkbook(ctx);
            using Stream data = file.OpenReadStream();
            return Results.Ok(svc.Save(WorkingEndpoints.ModelName(name), file.FileName, data, user.Identity?.Name ?? ""));
        });
        api.MapDelete("/models/{name}/instructions", (string name, WorkingInstructionService svc) =>
            svc.Delete(WorkingEndpoints.ModelName(name)) ? Results.NoContent() : Results.NotFound());
        api.MapGet("/models/{name}/instructions/file", (string name, WorkingInstructionService svc, HttpContext ctx) =>
        {
            if (svc.OpenFile(WorkingEndpoints.ModelName(name)) is not { } file) return Results.NotFound();
            ctx.Response.Headers.XContentTypeOptions = "nosniff";
            return Results.File(file.Data, XlsxType, file.FileName);
        });
        return api;
    }

    /// <summary>The "file" of a multipart upload; the form is read here (no IFormFile binding) with the limits raised to the workbook size.</summary>
    private static async Task<IFormFile> ReadWorkbook(HttpContext ctx)
    {
        if (!ctx.Request.HasFormContentType) throw new ArgumentException("Send the workbook as multipart/form-data.");
        const long totalLimit = WorkingInstructionService.MaxUploadBytes + 1024 * 1024;
        var limit = ctx.Features.Get<IHttpMaxRequestBodySizeFeature>();
        if (limit is { IsReadOnly: false }) limit.MaxRequestBodySize = totalLimit;
        ctx.Features.Set<IFormFeature>(new FormFeature(ctx.Request, new FormOptions { MultipartBodyLengthLimit = totalLimit }));
        IFormCollection form;
        try { form = await ctx.Request.ReadFormAsync(ctx.RequestAborted); }
        catch (InvalidDataException) { throw new ArgumentException("The workbook is too large (100 MB at most)."); }
        IFormFile? file = form.Files.GetFile("file");
        if (file is null || file.Length == 0) throw new ArgumentException("Choose the workbook to upload.");
        if (file.Length > WorkingInstructionService.MaxUploadBytes) throw new ArgumentException("The workbook is too large (100 MB at most).");
        return file;
    }
}
