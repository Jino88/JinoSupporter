using System.Windows;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Syncfusion.Blazor;
using BmesNgRateStandalone.Services;

namespace BmesNgRateStandalone;

public partial class App : Application
{
    public IServiceProvider Services { get; private set; } = default!;

    protected override void OnStartup(StartupEventArgs e)
    {
        var services = new ServiceCollection();
        var appPathsService = new AppPathsService();
        var appPaths = appPathsService.Current;
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["Database:Path"] = appPaths.MainDbPath,
                ["Schedule:Path"] = appPaths.ScheduleDbPath,
            })
            .Build();

        services.AddWpfBlazorWebView();
#if DEBUG
        services.AddBlazorWebViewDeveloperTools();
#endif

        // Syncfusion (Community license — works without a key)
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("");
        services.AddSyncfusionBlazor();

        // App services
        services.AddSingleton<IConfiguration>(configuration);
        services.AddSingleton(appPathsService);
        services.AddSingleton<WebRepository>();
        services.AddSingleton<NgRateSettingsService>();
        services.AddSingleton<NgRateService>();
        services.AddSingleton<NgRateReportService>();
        services.AddSingleton<BmesMaterialService>();
        services.AddSingleton<BmesRoutingScrapeService>();
        services.AddSingleton<BmesSettingsSyncService>();
        services.AddSingleton<WorkerStatusService>();
        services.AddSingleton<FCostService>();
        services.AddSingleton<FCostReportService>();

        Services = services.BuildServiceProvider();

        base.OnStartup(e);
    }
}
