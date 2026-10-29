using System.Security.Claims;
using System.Text.Json;
using Microsoft.AspNetCore.Http.Features;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

public static class TeamMemoEndpoints
{
    public sealed record FolderChange(string? ParentId, string Name, bool Plan);
    public sealed record MemoCreate(string? FolderId, string Title, string Kind);
    public sealed record HeaderChange(TeamMemo Original, TeamMemo Edited);
    public sealed record MemoChange(string? FolderId, string? Status, bool Pinned);
    public sealed record TextRequest(string Lang, string[] Texts);

    private static string User(ClaimsPrincipal user) => user.Identity?.Name ?? "";
    private static bool Own(ClaimsPrincipal user, string author) => user.IsInRole(AppRoles.Admin)
        || string.Equals(User(user), author, StringComparison.OrdinalIgnoreCase);
    private static bool HeaderDiffers(TeamMemo a, TeamMemo b) =>
        a.Title != b.Title || a.Kind != b.Kind || a.FolderId != b.FolderId || a.Owner != b.Owner || a.Tags != b.Tags
        || a.StartDate != b.StartDate || a.DueDate != b.DueDate || a.Description != b.Description
        || a.RepeatWeekdays != b.RepeatWeekdays || a.RepeatUntil != b.RepeatUntil;

    public static IEndpointRouteBuilder MapTeamMemo(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/portal/memo").RequireAuthorization();
        api.AddEndpointFilter(async (ctx, next) =>
        {
            if (!ctx.HttpContext.RequestServices.GetRequiredService<MenuPermissionService>()
                .IsAllowed(ctx.HttpContext.User, AppMenus.TeamMemo)) return Results.Forbid();
            if (!HttpMethods.IsGet(ctx.HttpContext.Request.Method)
                && ctx.HttpContext.Request.Headers["X-Requested-With"] != "PortalGNLAB")
                return Results.BadRequest(new { error = "Required request header is missing." });
            try { return await next(ctx); }
            catch (TeamMemoConflictException ex) { return Results.Conflict(new { error = ex.Message }); }
            catch (InvalidOperationException ex) { return Results.BadRequest(new { error = ex.Message }); }
            catch (ArgumentException ex) { return Results.BadRequest(new { error = ex.Message }); }
        });
        api.MapGet("", (TeamMemoService memos, WebRepository repo, TeamMemoAiSummaryService ai,
            TeamMemoTranslationService translations, ClaimsPrincipal user) => Results.Ok(new
        {
            folders = memos.GetFolders(), memos = memos.ListMemos(), currentUser = User(user), isAdmin = user.IsInRole(AppRoles.Admin),
            users = repo.GetAllUsers().Select(u => new { u.Username, displayName = string.IsNullOrWhiteSpace(u.DisplayName) ? u.Username : u.DisplayName }),
            aiConfigured = ai.IsConfigured, translationConfigured = translations.IsConfigured
        }));
        api.MapGet("/items/{id}", (string id, TeamMemoService memos) =>
            memos.GetMemo(id) is { } memo ? Results.Ok(memo) : Results.NotFound());
        api.MapPost("/items", (MemoCreate body, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            if (!TeamMemoKind.All.Contains(body.Kind)) return Results.BadRequest(new { error = "Unknown memo type." });
            return Results.Ok(new { id = memos.CreateMemo(body.FolderId, body.Title, body.Kind, User(user)) });
        });
        api.MapPut("/items/{id}/header", (string id, HeaderChange body, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            var fresh = memos.GetMemo(id);
            if (fresh is null) return Results.NotFound();
            if (HeaderDiffers(fresh, body.Original)) return Results.Conflict(new { error = "The memo header changed while you were editing. Revert to load it, then reapply your change." });
            if (!TeamMemoKind.All.Contains(body.Edited.Kind)) return Results.BadRequest(new { error = "Unknown memo type." });
            var edited = body.Edited;
            // Only header fields are written. Fresh status, pin and revision keep independent edits.
            fresh.Title = edited.Title; fresh.Kind = edited.Kind; fresh.FolderId = edited.FolderId; fresh.Owner = edited.Owner;
            fresh.Tags = edited.Tags; fresh.StartDate = edited.StartDate; fresh.DueDate = edited.DueDate;
            fresh.Description = edited.Description; fresh.RepeatWeekdays = edited.RepeatWeekdays; fresh.RepeatUntil = edited.RepeatUntil;
            memos.UpdateMemo(fresh, User(user)); return Results.NoContent();
        });
        api.MapPost("/items/{id}/{action}", (string id, string action, MemoChange body, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            if (memos.GetMemo(id) is null) return Results.NotFound();
            switch (action)
            {
                case "status":
                    if (!TeamMemoStatus.All.Contains(body.Status)) return Results.BadRequest(new { error = "Unknown status." });
                    memos.SetMemoStatus(id, body.Status!, User(user)); break;
                case "pin": memos.SetPinned(id, body.Pinned, User(user)); break;
                case "move": memos.MoveMemo(id, body.FolderId, User(user)); break;
                default: return Results.NotFound();
            }
            return Results.NoContent();
        });
        api.MapDelete("/items/{id}", (string id, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            var memo = memos.GetMemo(id); if (memo is null) return Results.NotFound();
            if (!Own(user, memo.CreatedBy) && !Own(user, memo.Owner)) return Results.Forbid();
            memos.DeleteMemo(id, User(user)); return Results.NoContent();
        });
        api.MapPost("/folders", (FolderChange body, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            var id = memos.AddFolder(body.ParentId, body.Name, User(user));
            if (body.Plan) memos.SetFolderPlan(id, true);
            return Results.NoContent();
        });
        api.MapPut("/folders/{id}/{action}", (string id, string action, FolderChange body, TeamMemoService memos) =>
        {
            switch (action)
            {
                case "rename": memos.RenameFolder(id, body.Name); break;
                case "move": memos.MoveFolder(id, body.ParentId); break;
                case "plan": memos.SetFolderPlan(id, body.Plan); break;
                default: return Results.NotFound();
            }
            return Results.NoContent();
        });
        api.MapDelete("/folders/{id}", (string id, TeamMemoService memos) => { memos.DeleteFolder(id); return Results.NoContent(); });
        api.MapPost("/plans/initialize", (TeamMemoService memos, WebRepository repo, ClaimsPrincipal user) =>
        {
            var folders = memos.GetFolders();
            if (repo.GetSetting("TeamMemo:PlanTypesSeeded") == "1" && TeamMemoService.PlanFolderIds(folders).Count > 0)
                return Results.NoContent();
            TeamMemoFolder? TopLevel(string name) => folders.FirstOrDefault(f => f.ParentId is null && string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase));
            if (TopLevel("Plan") is { IsPlan: true } old && TopLevel("Other") is null) memos.RenameFolder(old.Id, "Other");
            folders = memos.GetFolders();
            foreach (string type in new[] { "Personal", "Work", "New Product", "Other" })
                memos.SetFolderPlan(TopLevel(type)?.Id ?? memos.AddFolder(null, type, User(user)), true);
            repo.SetSetting("TeamMemo:PlanTypesSeeded", "1");
            return Results.NoContent();
        });

        // Multipart is read explicitly; the group filter enforces the same custom-header CSRF guard.
        api.MapPost("/items/{id}/entries", async (string id, HttpContext ctx, TeamMemoService memos) =>
        {
            if (memos.GetMemo(id) is null) return Results.NotFound();
            var (form, files) = await Upload(ctx);
            var entry = JsonSerializer.Deserialize<TeamMemoEntry>(form["entry"].ToString(), new JsonSerializerOptions(JsonSerializerDefaults.Web))
                ?? throw new InvalidOperationException("Entry is required.");
            if (string.IsNullOrWhiteSpace(entry.Title)) return Results.BadRequest(new { error = "Entry title is required." });
            entry.MemoId = id;
            if (entry.Id.Length == 0) memos.AddEntry(id, entry.EntryDate, entry.Title, entry.Body, entry.TargetDate, User(ctx.User), files);
            else memos.UpdateEntry(entry, User(ctx.User), files);
            return Results.NoContent();
        });
        api.MapPut("/items/{id}/entry-done", (string id, TeamMemoEntry body, TeamMemoService memos, ClaimsPrincipal user) =>
        { memos.SetEntryDone(id, body.Id, body.DoneDate, User(user)); return Results.NoContent(); });
        api.MapPost("/items/{id}/entry-delete", (string id, TeamMemoEntry body, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            var entry = memos.GetMemo(id)?.Entries.FirstOrDefault(e => e.Id == body.Id);
            if (entry is null) return Results.NotFound();
            if (!Own(user, entry.Author)) return Results.Forbid();
            memos.DeleteEntry(id, body.Id, User(user)); return Results.NoContent();
        });
        api.MapPost("/items/{id}/comments", async (string id, HttpContext ctx, TeamMemoService memos) =>
        {
            var memo = memos.GetMemo(id); if (memo is null) return Results.NotFound();
            var (form, files) = await Upload(ctx);
            string? entryId = string.IsNullOrEmpty(form["entryId"]) ? null : form["entryId"].ToString();
            if (entryId is not null && !memo.Entries.Any(e => e.Id == entryId)) return Results.NotFound();
            if (string.IsNullOrWhiteSpace(form["body"]) && files.Count == 0) return Results.BadRequest(new { error = "Write a comment or attach a file." });
            memos.AddComment(id, entryId, form["body"].ToString(), form["isResult"] == "true", User(ctx.User), files);
            return Results.NoContent();
        });
        api.MapPost("/items/{id}/comments/{commentId:long}/{action}", (string id, long commentId, string action,
            TeamMemoComment body, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            var memo = memos.GetMemo(id); if (memo is null) return Results.NotFound();
            var comment = memo.Comments.Concat(memo.Entries.SelectMany(e => e.Comments)).FirstOrDefault(c => c.Id == commentId);
            if (comment is null) return Results.NotFound();
            if (!Own(user, comment.Author) && !(action == "result" && Own(user, memo.Owner))) return Results.Forbid();
            switch (action)
            {
                case "edit": memos.UpdateComment(id, commentId, body.Body.Trim()); break;
                case "result": memos.SetCommentResult(id, commentId, body.IsResult); break;
                case "delete": memos.DeleteComment(id, commentId); break;
                default: return Results.NotFound();
            }
            return Results.NoContent();
        });
        api.MapGet("/attachments/{id}", (string id, bool? download, TeamMemoService memos, HttpContext ctx) =>
        {
            var file = memos.GetAttachment(id); if (file is null) return Results.NotFound();
            ctx.Response.Headers.XContentTypeOptions = "nosniff";
            // Inline only bitmap images; HTML/SVG and other active content is always downloaded.
            bool inline = download != true && new[] { "image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp" }.Contains(file.Value.ContentType.ToLowerInvariant());
            return Results.File(file.Value.Data, file.Value.ContentType, inline ? null : file.Value.FileName);
        });
        api.MapDelete("/items/{id}/attachments/{fileId}", (string id, string fileId, TeamMemoService memos, ClaimsPrincipal user) =>
        {
            var memo = memos.GetMemo(id); if (memo is null) return Results.NotFound();
            var file = memo.Attachments.Concat(memo.Entries.SelectMany(e => e.Attachments))
                .Concat(memo.Comments.Concat(memo.Entries.SelectMany(e => e.Comments)).SelectMany(c => c.Attachments)).FirstOrDefault(f => f.Id == fileId);
            if (file is null) return Results.NotFound();
            if (!Own(user, file.Author)) return Results.Forbid();
            memos.DeleteAttachment(id, fileId, User(user)); return Results.NoContent();
        });
        api.MapGet("/items/{id}/ai-status", (string id, TeamMemoService memos) => Results.Ok(new
        {
            runningBy = TeamMemoAiSummaryService.RunningBy(id),
            entries = memos.GetMemo(id)?.Entries.ToDictionary(e => e.Id, e => TeamMemoAiSummaryService.RunningByEntry(id, e.Id))
        }));
        api.MapPost("/items/{id}/summary", async (string id, TeamMemoEntry body, TeamMemoAiSummaryService ai, ClaimsPrincipal user) =>
        {
            // Paid work is saved even when the requesting page closes, matching Blazor.
            if (string.IsNullOrEmpty(body.Id)) await ai.SummarizeAsync(id, User(user), CancellationToken.None);
            else await ai.SummarizeEntryAsync(id, body.Id, User(user), CancellationToken.None);
            return Results.NoContent();
        });
        api.MapPost("/translations/cached", (TextRequest body, TeamMemoTranslationService translations) =>
            ValidLanguage(body.Lang) ? Results.Ok(translations.Cached(body.Texts, body.Lang)) : Results.BadRequest(new { error = "Unknown language." }));
        api.MapPost("/translations/fetch", async (TextRequest body, TeamMemoTranslationService translations) =>
            ValidLanguage(body.Lang) ? Results.Ok(await translations.TranslateAsync(body.Texts, body.Lang)) : Results.BadRequest(new { error = "Unknown language." }));
        api.MapGet("/translations/status", () => Results.Ok(new { running = TeamMemoTranslationService.AllRunning,
            done = TeamMemoTranslationService.AllDone, total = TeamMemoTranslationService.AllTotal, message = TeamMemoTranslationService.AllMessage }));
        api.MapPost("/translations/all", async (TeamMemoTranslationService translations) =>
        { await translations.TranslateAllMemosAsync(); return Results.NoContent(); });
        return app;
    }

    private static bool ValidLanguage(string lang) => TeamMemoLanguage.Targets.Contains(lang);
    private static async Task<(IFormCollection Form, List<(string Name, string ContentType, byte[] Data)> Files)> Upload(HttpContext ctx)
    {
        const long totalLimit = 20 * TeamMemoService.MaxAttachmentBytes + 1024 * 1024;
        var limit = ctx.Features.Get<IHttpMaxRequestBodySizeFeature>();
        if (limit is { IsReadOnly: false }) limit.MaxRequestBodySize = totalLimit;
        ctx.Features.Set<IFormFeature>(new FormFeature(ctx.Request, new FormOptions { MultipartBodyLengthLimit = totalLimit }));
        var form = await ctx.Request.ReadFormAsync(ctx.RequestAborted);
        if (form.Files.Count > 20) throw new InvalidOperationException("Select at most 20 files.");
        var files = new List<(string, string, byte[])>();
        foreach (var file in form.Files)
        {
            if (file.Length > TeamMemoService.MaxAttachmentBytes) throw new InvalidOperationException($"{file.FileName} is larger than 20 MB.");
            using var data = new MemoryStream(); await file.CopyToAsync(data, ctx.RequestAborted);
            files.Add((Path.GetFileName(file.FileName), file.ContentType, data.ToArray()));
        }
        return (form, files);
    }
}
