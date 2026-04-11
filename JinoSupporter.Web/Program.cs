using System.Security.Claims;
using JinoSupporter.Web.Components;
using JinoSupporter.Web.Services;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Components.Server.Circuits;

var builder = WebApplication.CreateBuilder(args);

// ── Services ──────────────────────────────────────────────────────────────────

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddHubOptions(options =>
    {
        options.MaximumReceiveMessageSize = 50 * 1024 * 1024; // 50 MB
    });

// Cookie authentication
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(o =>
    {
        o.LoginPath         = "/login";
        o.AccessDeniedPath  = "/login";
        o.ExpireTimeSpan    = TimeSpan.FromDays(30);
        o.SlidingExpiration = true;
    });
builder.Services.AddAuthorizationCore();
builder.Services.AddCascadingAuthenticationState();

// Claude HTTP client
builder.Services.AddHttpClient<ClaudeService>(client =>
{
    client.BaseAddress = new Uri("https://api.anthropic.com/v1/");
    client.Timeout     = TimeSpan.FromSeconds(180);
});

// Singleton DB repository (SQLite file shared across all requests)
builder.Services.AddSingleton<WebRepository>();

// Connected-users tracking (singleton service + scoped circuit handler)
builder.Services.AddSingleton<ConnectedUsersService>();
builder.Services.AddScoped<UserCircuitHandler>();
builder.Services.AddScoped<CircuitHandler>(sp => sp.GetRequiredService<UserCircuitHandler>());

// Listen on all interfaces
builder.WebHost.UseUrls("http://*:5050");

// ── Pipeline ──────────────────────────────────────────────────────────────────

var app = builder.Build();

// Seed default admin user if no users exist
var repo = app.Services.GetRequiredService<WebRepository>();
if (repo.GetAllUsers().Count == 0)
    repo.AddUser("admin", AuthService.HashPassword("admin123"), AppRoles.Admin);

if (!app.Environment.IsDevelopment())
    app.UseExceptionHandler("/Error", createScopeForErrors: true);

app.UseStaticFiles();
app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();

// ── Auth endpoints ────────────────────────────────────────────────────────────

app.MapPost("/auth/login", async (HttpContext ctx) =>
{
    var form     = await ctx.Request.ReadFormAsync();
    string user  = form["username"].ToString().Trim();
    string pass  = form["password"].ToString();
    string ret   = form["returnUrl"].ToString();
    if (string.IsNullOrWhiteSpace(ret) || !ret.StartsWith('/')) ret = "/";

    var record = repo.GetUser(user);
    if (record is null || !AuthService.VerifyPassword(pass, record.PasswordHash))
    {
        ctx.Response.Redirect($"/login?error=1&returnUrl={Uri.EscapeDataString(ret)}");
        return;
    }

    var claims    = new[] { new Claim(ClaimTypes.Name, record.Username), new Claim(ClaimTypes.Role, record.Role) };
    var identity  = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
    var principal = new ClaimsPrincipal(identity);
    await ctx.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal);
    ctx.Response.Redirect(ret);
}).DisableAntiforgery();

app.MapPost("/auth/logout", async (HttpContext ctx) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    ctx.Response.Redirect("/login");
}).DisableAntiforgery();

// ── Blazor ────────────────────────────────────────────────────────────────────

app.MapRazorComponents<App>()
   .AddInteractiveServerRenderMode();

app.Run();
