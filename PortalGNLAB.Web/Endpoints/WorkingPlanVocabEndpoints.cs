using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>Working › Dev Plan standard names: GET for every Working user, PUT (Admin only) saves the cleaned-up list.</summary>
public static class WorkingPlanVocabEndpoints
{
    public static IEndpointRouteBuilder MapWorkingPlanVocab(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/portal/tools/working/plan-vocab").RequireAuthorization();
        api.AddEndpointFilter(async (ctx, next) =>
        {
            if (!ctx.HttpContext.RequestServices.GetRequiredService<MenuPermissionService>()
                .IsAllowed(ctx.HttpContext.User, AppMenus.Working)) return Results.Forbid();
            if (!HttpMethods.IsGet(ctx.HttpContext.Request.Method)
                && ctx.HttpContext.Request.Headers["X-Requested-With"] != "PortalGNLAB")
                return Results.BadRequest(new { error = "Required request header is missing." });
            try { return await next(ctx); }
            catch (ArgumentException ex) { return Results.BadRequest(new { error = ex.Message }); }
        });

        api.MapGet("", (WorkingPlanVocabService svc) => Results.Ok(new { vocab = svc.Get() }));
        api.MapPut("", async (HttpContext http, WorkingPlanVocabService svc) =>
        {
            if (!http.User.IsInRole(AppRoles.Admin))
                return Results.Json(new { error = "Only an administrator can change the standard names." }, statusCode: StatusCodes.Status403Forbidden);
            if (http.Request.ContentLength > WorkingPlanVocabService.MaxBytes)
                return Results.BadRequest(new { error = "The list is too large." });
            // read at most the limit + 1 byte, so a body without a length header is cut off too
            var buffer = new MemoryStream();
            var chunk = new byte[8192];
            int n;
            while ((n = await http.Request.Body.ReadAsync(chunk)) > 0)
            {
                buffer.Write(chunk, 0, n);
                if (buffer.Length > WorkingPlanVocabService.MaxBytes)
                    return Results.BadRequest(new { error = "The list is too large." });
            }
            string text = System.Text.Encoding.UTF8.GetString(buffer.ToArray());
            // the body is {vocab: {...}}; the service takes the inner object
            string inner;
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(text);
                if (doc.RootElement.ValueKind != System.Text.Json.JsonValueKind.Object
                    || !doc.RootElement.TryGetProperty("vocab", out var v)) throw new ArgumentException("vocab is required.");
                inner = v.GetRawText();
            }
            catch (System.Text.Json.JsonException) { throw new ArgumentException("The list is not valid JSON."); }
            return Results.Ok(new { vocab = svc.Save(inner) });
        });
        return api;
    }
}
