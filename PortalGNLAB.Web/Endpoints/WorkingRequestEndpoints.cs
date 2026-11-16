using System.Security.Claims;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>Working > Request (JIG / MC PROCESSING REQUEST): mapped inside the Working API group (auth and header checks there).</summary>
public static class WorkingRequestEndpoints
{
    public static RouteGroupBuilder MapWorkingRequests(this RouteGroupBuilder api)
    {
        api.MapGet("/requests", (WorkingRequestService svc) => Results.Ok(new { requests = svc.List() }));
        api.MapPost("/requests", (WorkingRequestSave body, WorkingRequestService svc, ClaimsPrincipal user) =>
            Results.Ok(svc.Create(body, user.Identity?.Name ?? "")));
        api.MapPut("/requests/{id}", (string id, WorkingRequestSave body, WorkingRequestService svc, ClaimsPrincipal user) =>
            svc.Update(id, body, user.Identity?.Name ?? "") is { } r ? Results.Ok(r) : Results.NotFound());
        api.MapDelete("/requests/{id}", (string id, WorkingRequestService svc) =>
            svc.Delete(id) ? Results.NoContent() : Results.NotFound());
        // AI input: the form filled from pasted text / pictures; nothing is saved here
        api.MapPost("/requests/ai", async (RequestAiRequest body, WorkingRequestAiService ai, HttpContext http) =>
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(http.RequestAborted);
            cts.CancelAfter(TimeSpan.FromMinutes(3));
            var result = await ai.FillAsync(body, cts.Token);
            return Results.Ok(new { request = result.Request, note = result.Note, model = result.Model });
        }).WithMetadata(new Microsoft.AspNetCore.Mvc.RequestSizeLimitAttribute(40L * 1024 * 1024)); // up to 6 pasted pictures as base64
        // Drawing No lookup in the JIG/MC drawing DB (read-only); never throws, a missing DB comes back as ok=false
        api.MapPost("/requests/drawings/check", (DrawingCheckBody body, JigDrawingDbService db) =>
        {
            var r = db.Check((body.DrawingNos ?? []).Take(JigDrawingDbService.MaxCheck));
            return Results.Ok(new { ok = r.Ok, error = r.Error, path = r.Path, results = r.Results });
        });
        api.MapPost("/requests/drawings/suggest", (DrawingSuggestBody body, JigDrawingDbService db) =>
        {
            var r = db.Suggest(body.Name, body.Model, body.Limit ?? 10);
            return Results.Ok(new { ok = r.Ok, error = r.Error, path = r.Path, suggestions = r.Suggestions });
        });
        return api;
    }

    public sealed record DrawingCheckBody(string[]? DrawingNos);

    public sealed record DrawingSuggestBody(string? Name, string? Model, int? Limit);
}
