using System.Security.Claims;
using Microsoft.AspNetCore.Http.Features;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>Working API: model groups (read-only), per-model info and dev / issue / review entries.</summary>
public static class WorkingEndpoints
{
    private static string User(ClaimsPrincipal user) => user.Identity?.Name ?? "";

    /// <summary>
    /// The model name of a {name} route value. The server decodes the path except %2F, so a model with a '/' in its name
    /// (MSM-X526/X626B) arrives as MSM-X526%2FX626B and would be kept under that name.
    /// </summary>
    public static string ModelName(string name) => name.Replace("%2F", "/", StringComparison.OrdinalIgnoreCase);

    /// <summary>The reserved model-name prefix of a personal plan: the model name is ~personal:&lt;login id&gt;.</summary>
    public const string PersonalPrefix = "~personal:";

    public static bool IsPersonalModel(string? modelName) =>
        modelName is not null && modelName.StartsWith(PersonalPrefix, StringComparison.OrdinalIgnoreCase);

    public static string PersonalKey(string user) => PersonalPrefix + user;

    /// <summary>True for an ordinary model, or for a personal key that belongs to <paramref name="user"/> (case-insensitive).</summary>
    public static bool PersonalOwnerOk(string? modelName, string? user)
    {
        if (!IsPersonalModel(modelName)) return true;
        string owner = modelName![PersonalPrefix.Length..];
        return !string.IsNullOrEmpty(user) && owner.Length > 0 && string.Equals(owner, user, StringComparison.OrdinalIgnoreCase);
    }

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
            // a personal plan (~personal:<id>) is only for its owner, on every /models/{name}/** path
            if (ctx.HttpContext.Request.RouteValues["name"] is string routeName
                && !PersonalOwnerOk(ModelName(routeName), ctx.HttpContext.User.Identity?.Name))
                return Results.StatusCode(StatusCodes.Status403Forbidden);
            try { return await next(ctx); }
            catch (WorkingConflictException ex) { return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status409Conflict); }
            catch (ArgumentException ex) { return Results.BadRequest(new { error = ex.Message }); }
        });

        // English / Korean / Vietnamese: cached translations at once, what is missing is translated in the background
        api.MapPost("/translations", (WorkingTranslationRequestBody body, WorkingTranslationService tr) =>
            Results.Ok(tr.Translations(body.Lang, body.Texts)));
        api.MapGet("/translations/status", (WorkingTranslationService tr) =>
        {
            var s = tr.Status();
            return Results.Ok(new { running = s.Running, done = s.Done, total = s.Total, message = s.Message });
        });

        api.MapGet("/groups", (WorkingService svc) => Results.Ok(svc.ListGroups()));
        // what the tree hides / adds (kept in PortalGNLAB.db; the Glossary DB is never written)
        api.MapGet("/tree", (WorkingService svc) => Results.Ok(svc.GetTree()));
        api.MapPut("/tree/hidden", (WorkingHiddenSave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.SetHidden(body.Hidden, User(user))));
        api.MapPost("/tree/models", (WorkingTreeModelSave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.AddTreeModel(body.Name, body.GroupId, User(user))));
        api.MapDelete("/tree/models/{name}", (string name, WorkingService svc) =>
            svc.DeleteTreeModel(ModelName(name)) ? Results.NoContent() : Results.NotFound());
        api.MapPut("/tree/models/{name}/group", (string name, WorkingTreeMoveSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.MoveTreeModel(ModelName(name), body.GroupId, User(user)) is { } moved ? Results.Ok(moved) : Results.NotFound());
        // sub-groups made in Manage
        api.MapPost("/tree/groups", (WorkingTreeGroupSave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.AddTreeGroup(body.Name, body.ParentId, User(user))));
        api.MapPut("/tree/groups/{id}", (string id, WorkingTreeGroupSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.RenameTreeGroup(id, body.Name, User(user)) is { } g ? Results.Ok(g) : Results.NotFound());
        api.MapDelete("/tree/groups/{id}", (string id, WorkingService svc) =>
            svc.DeleteTreeGroup(id) ? Results.NoContent() : Results.NotFound());

        api.MapGet("/models/{name}", (string name, WorkingService svc) => Results.Ok(svc.GetModel(ModelName(name))));
        // hand-entered process table: AI turns a free-form request into the changed table (nothing is saved here)
        api.MapPost("/models/{name}/manual/ai", async (string name, ManualAiRequest body, WorkingManualAiService ai, HttpContext http) =>
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(http.RequestAborted);
            cts.CancelAfter(TimeSpan.FromMinutes(3));
            return Results.Ok(await ai.EditAsync(ModelName(name), body, cts.Token));
        });
        // AI tab: AI decides where free text belongs (Model Info / Dev Plan / entries); nothing is saved here
        api.MapPost("/models/{name}/ai/place", async (string name, PlaceAiRequest body, WorkingPlaceAiService ai, HttpContext http) =>
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(http.RequestAborted);
            cts.CancelAfter(TimeSpan.FromMinutes(3));
            return Results.Ok(await ai.PlaceAsync(ModelName(name), body, cts.Token));
        }).WithMetadata(new Microsoft.AspNetCore.Mvc.RequestSizeLimitAttribute(40L * 1024 * 1024)); // up to 6 pasted pictures as base64
        api.MapPut("/models/{name}", (string name, WorkingModelSave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.SaveModel(ModelName(name), body, User(user))));

        api.MapGet("/models/{name}/entries", (string name, WorkingService svc) => Results.Ok(new { entries = svc.ListEntries(ModelName(name)) }));
        api.MapPost("/models/{name}/entries", (string name, WorkingEntrySave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.CreateEntry(ModelName(name), body, User(user))));
        api.MapPut("/models/{name}/entries/{id}", (string name, string id, WorkingEntrySave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.UpdateEntry(ModelName(name), id, body, User(user)) is { } e ? Results.Ok(e) : Results.NotFound());
        api.MapDelete("/models/{name}/entries/{id}", (string name, string id, WorkingService svc) =>
            svc.DeleteEntry(ModelName(name), id) ? Results.NoContent() : Results.NotFound());

        // dated progress notes under an entry (issues)
        api.MapPost("/models/{name}/entries/{id}/logs", (string name, string id, WorkingLogSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.CreateLog(ModelName(name), id, body, User(user)) is { } l ? Results.Ok(l) : Results.NotFound());
        api.MapPut("/models/{name}/entries/{id}/logs/{logId}", (string name, string id, string logId, WorkingLogSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.UpdateLog(ModelName(name), id, logId, body, User(user)) is { } l ? Results.Ok(l) : Results.NotFound());
        api.MapDelete("/models/{name}/entries/{id}/logs/{logId}", (string name, string id, string logId, WorkingService svc) =>
            svc.DeleteLog(ModelName(name), id, logId) ? Results.NoContent() : Results.NotFound());

        // Setup Plan rows; their problems are issues with a SetupRowId
        api.MapGet("/models/{name}/setup", (string name, WorkingService svc) => Results.Ok(new { rows = svc.ListSetupRows(ModelName(name)) }));
        api.MapPost("/models/{name}/setup", (string name, WorkingSetupRowSave body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.CreateSetupRow(ModelName(name), body, User(user))));
        // Plan > AI input: AI turns free text into Setup Plan rows; nothing is saved here
        api.MapPost("/models/{name}/setup/ai", async (string name, SetupAiRequest body, WorkingSetupAiService ai, HttpContext http) =>
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(http.RequestAborted);
            cts.CancelAfter(TimeSpan.FromMinutes(3));
            return Results.Ok(await ai.SetupAsync(ModelName(name), body, cts.Token));
        }).WithMetadata(new Microsoft.AspNetCore.Mvc.RequestSizeLimitAttribute(40L * 1024 * 1024)); // up to 6 pasted pictures as base64
        api.MapPost("/models/{name}/setup/import", (string name, WorkingSetupImport body, WorkingService svc, ClaimsPrincipal user) =>
            Results.Ok(new { rows = svc.ImportSetupRows(ModelName(name), body.Rows ?? [], User(user)) }));
        api.MapPut("/models/{name}/setup/{rowId}", (string name, string rowId, WorkingSetupRowSave body, WorkingService svc, ClaimsPrincipal user) =>
            svc.UpdateSetupRow(ModelName(name), rowId, body, User(user)) is { } r ? Results.Ok(r) : Results.NotFound());
        api.MapDelete("/models/{name}/setup/{rowId}", (string name, string rowId, WorkingService svc) =>
            svc.DeleteSetupRow(ModelName(name), rowId) ? Results.NoContent() : Results.NotFound());

        // pictures and videos pasted into an entry (no logId) or into one of its notes, stored as sent
        api.MapPost("/models/{name}/entries/{id}/images", async (string name, string id, HttpContext ctx, WorkingService svc, ClaimsPrincipal user) =>
        {
            (string? logId, List<WorkingUpload> files) = await ReadPictures(ctx);
            return svc.AddFiles(ModelName(name), id, logId, files, User(user)) is { } images ? Results.Ok(new { images }) : Results.NotFound();
        });
        // streamed from the DB with byte ranges, so a video can be played from any point
        api.MapGet("/images/{imageId}", (string imageId, WorkingService svc, HttpContext ctx) =>
        {
            // a picture of someone else's personal plan is not there for this user
            if (!PersonalOwnerOk(svc.ImageModelName(imageId), ctx.User.Identity?.Name)) return Results.NotFound();
            if (svc.OpenImage(imageId) is not { } file) return Results.NotFound();
            ctx.Response.RegisterForDispose(file);
            ctx.Response.Headers.XContentTypeOptions = "nosniff";
            return Results.Stream(file.Data, file.ContentType, enableRangeProcessing: true);
        });
        api.MapDelete("/models/{name}/entries/{id}/images/{imageId}", (string name, string id, string imageId, WorkingService svc) =>
            svc.DeleteImage(ModelName(name), id, imageId) ? Results.NoContent() : Results.NotFound());

        // BMES model groups and a product's BOM for Model Info (read-only on the BMES side)
        api.MapGet("/bmes/groups", (WebRepository repo) =>
            Results.Ok(new { materials = WorkingBmesService.BuildMaterials(repo.GetModelGroups()) }));
        api.MapGet("/bmes/bom", (string? product, BmesBomCacheService cache) =>
        {
            string name = (product ?? "").Trim();
            if (name.Length == 0) return Results.BadRequest(new { error = "product is required." });
            return Results.Ok(cache.TryLoadLatest(name) is { } entry
                ? WorkingBmesService.ToBom(name, "cache", entry)
                : WorkingBmesService.Empty(name));
        });
        // read from the BMES BOM screen (MES071050), the same tree the screen shows
        api.MapPost("/bmes/bom", async (WorkingBmesBomRequest body, NgRateSettingsService settings,
            BmesBomCacheService cache, CancellationToken ct) =>
        {
            string name = (body.Product ?? "").Trim();
            if (name.Length == 0) return Results.BadRequest(new { error = "product is required." });
            if (!settings.IsCredentialsConfigured)
                return Results.Ok(WorkingBmesService.Empty(name, "The BMES login is not configured in Setting."));
            try
            {
                using var session = await BmesGateway.OpenBomSessionAsync(settings, ct);
                var (rows, error) = await BmesGateway.FetchBomAsync(session, WorkingBmesService.Plant, name, cancellationToken: ct);
                if (rows is null) return Results.Ok(WorkingBmesService.Empty(name, error));
                var query = new BmesBomCacheQuery(name, WorkingBmesService.Plant, DateTime.Today,
                    WorkingBmesService.MaxDepth, WorkingBmesService.MaxRows);
                cache.Save(query, rows);
                return Results.Ok(cache.TryLoad(query) is { } saved
                    ? WorkingBmesService.ToBom(name, "server", saved)
                    : WorkingBmesService.Empty(name));
            }
            catch (Exception ex)
            {
                return Results.Ok(WorkingBmesService.Empty(name, "BOM could not be read from the BMES server: " + ex.Message));
            }
        });

        // does a BMES model of this name exist: catalog, model groups and (when connected) the BMES product master
        api.MapGet("/bmes/check", async (string? name, BmesBomCacheService cache, WebRepository repo,
            BmesFcostActualService fcostActual, AppPathsService paths) =>
        {
            string text = (name ?? "").Trim();
            if (text.Length == 0) return Results.BadRequest(new { error = "name is required." });
            var catalog = WorkingBmesService.CatalogCandidates(text, cache.LoadModelCatalog(WorkingBmesService.Plant).Models);
            var group = WorkingBmesService.GroupCandidates(text, repo.GetModelGroups());
            var bmes = new List<WorkingModelCandidate>();
            bool checkedBmes = false;
            string bmesError = "";
            var connection = BmesReportEndpoints.BuildNetworkDbConnection(paths);
            if (!string.IsNullOrWhiteSpace(connection.Server) && !string.IsNullOrWhiteSpace(connection.UserId))
            {
                try
                {
                    var found = await fcostActual.SearchBomModelsAsync(new BmesBomModelSearchQuery
                    {
                        Connection = connection,
                        SearchText = text,
                        Plant = WorkingBmesService.Plant,
                        WorkDate = DateTime.Today,
                        MaxRows = WorkingBmesService.MaxCandidates,
                    });
                    bmes.AddRange(found.Select(m => new WorkingModelCandidate(m.ProductName.Trim(), m.ProductCode.Trim(), "bmes")));
                    checkedBmes = true;
                }
                catch (Exception ex) { bmesError = ex.Message; }
            }
            var (exact, candidates) = WorkingBmesService.MergeCandidates(text, bmes, catalog, group);
            return Results.Ok(new { name = text, exact, candidates, bmesChecked = checkedBmes, bmesError });
        });

        // PDM drawings for the BOM materials (same lookups as BOM & Drawing; the BMES side is read-only)
        api.MapPost("/bmes/pdm", (WorkingPdmRequest body, BmesPdmScrapeService pdm) =>
        {
            var names = (body.Names ?? []).Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
            if (names.Count == 0) return Results.Ok(new { rows = Array.Empty<object>() });
            var matches = pdm.SearchByNames(names, []);
            return Results.Ok(new { rows = matches.Select(m => new { m.Pdmno, m.Pdmvr, m.PdtnoTx, m.Pdrtx, m.Erdat, m.Ardat }).ToList() });
        });
        api.MapGet("/bmes/drawing", async (string? pdmno, string? pdmvr, BmesPdmScrapeService pdm, HttpContext ctx) =>
        {
            string no = (pdmno ?? "").Trim();
            if (no.Length == 0) return Results.BadRequest(new { error = "pdmno is required." });
            var result = await pdm.DownloadDrawingAsync(no, string.IsNullOrWhiteSpace(pdmvr) ? "0" : pdmvr.Trim());
            if (!result.Success) return Results.BadRequest(new { error = result.Error });
            var disposition = new Microsoft.Net.Http.Headers.ContentDispositionHeaderValue("inline");
            disposition.SetHttpFileName(result.FileName);
            ctx.Response.Headers.ContentDisposition = disposition.ToString();
            ctx.Response.Headers.XContentTypeOptions = "nosniff";
            return Results.File(result.Bytes, WorkingBmesService.DrawingContentType(result.FileName));
        });

        return app;
    }

    public sealed record WorkingBmesBomRequest(string? Product);
    public sealed record WorkingHiddenSave(List<string?>? Hidden);
    public sealed record WorkingTranslationRequestBody(string? Lang, List<string?>? Texts);
    public sealed record WorkingTreeModelSave(string? Name, string? GroupId);
    public sealed record WorkingTreeMoveSave(string? GroupId);
    public sealed record WorkingTreeGroupSave(string? Name, string? ParentId);
    public sealed record WorkingPdmRequest(List<string>? Names);

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
