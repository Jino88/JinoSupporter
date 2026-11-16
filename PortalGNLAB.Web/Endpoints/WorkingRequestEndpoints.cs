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
        return api;
    }
}
