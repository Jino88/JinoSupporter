using System.Security.Claims;
using Microsoft.AspNetCore.Http.Features;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>Working > Process List (공정편성 LIST): mapped inside the Working API group (auth and header checks there).</summary>
public static class WorkingProcessListEndpoints
{
    public sealed record ProcessListSave(List<WorkingProcessRow?>? Rows, int Revision);

    public static RouteGroupBuilder MapWorkingProcessList(this RouteGroupBuilder api)
    {
        api.MapGet("/models/{name}/process-list", (string name, WorkingProcessListService svc) =>
            Results.Ok(svc.Get(WorkingEndpoints.ModelName(name))));
        api.MapPut("/models/{name}/process-list", (string name, ProcessListSave body, WorkingProcessListService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.Save(WorkingEndpoints.ModelName(name), body.Rows ?? [], body.Revision, user.Identity?.Name ?? "")));
        // a JIG / Machine / Tester list workbook read into the rows, in place of what is there
        api.MapPost("/models/{name}/process-list/import", async (string name, HttpContext ctx, WorkingProcessListService svc, ClaimsPrincipal user) =>
        {
            (IFormFile file, int revision) = await ReadWorkbook(ctx);
            await using Stream s = file.OpenReadStream();
            List<WorkingProcessRow> rows = WorkingProcessListService.ReadWorkbook(s);
            return Results.Ok(svc.Save(WorkingEndpoints.ModelName(name), rows, revision, user.Identity?.Name ?? "", Path.GetFileName(file.FileName)));
        });
        return api;
    }

    private static async Task<(IFormFile File, int Revision)> ReadWorkbook(HttpContext ctx)
    {
        if (!ctx.Request.HasFormContentType) throw new ArgumentException("Send the workbook as multipart/form-data.");
        const long totalLimit = WorkingProcessListService.MaxUploadBytes + 1024 * 1024;
        var limit = ctx.Features.Get<IHttpMaxRequestBodySizeFeature>();
        if (limit is { IsReadOnly: false }) limit.MaxRequestBodySize = totalLimit;
        ctx.Features.Set<IFormFeature>(new FormFeature(ctx.Request, new FormOptions { MultipartBodyLengthLimit = totalLimit }));
        IFormCollection form;
        try { form = await ctx.Request.ReadFormAsync(ctx.RequestAborted); }
        catch (InvalidDataException) { throw new ArgumentException("The workbook is too large (50 MB at most)."); }
        IFormFile file = form.Files.GetFile("file") ?? throw new ArgumentException("Choose an Excel workbook.");
        string ext = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (ext is not (".xlsx" or ".xlsm")) throw new ArgumentException("Upload the JIG / Machine list as an .xlsx workbook.");
        return (file, int.TryParse(form["revision"], out int revision) ? revision : -1);
    }
}
