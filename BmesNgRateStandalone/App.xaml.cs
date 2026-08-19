using System.Windows;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using BmesNgRateStandalone.Services;
using BmesNgRateStandalone.Services.BmesReports;

namespace BmesNgRateStandalone;

public partial class App : Application
{
    public IServiceProvider Services { get; private set; } = default!;

    protected override void OnStartup(StartupEventArgs e)
    {
        InstallGlobalExceptionHandlers();
        StandaloneErrorLog.Write("Startup", $"Starting BMES NG Rate Standalone {GetType().Assembly.GetName().Version}");

        try
        {
            var services = new ServiceCollection();
            var appPathsService = new AppPathsService();
            var appPaths = appPathsService.Current;
            var configuration = new ConfigurationBuilder()
                .AddInMemoryCollection(new Dictionary<string, string?>
                {
                    ["Database:Path"] = appPaths.MainDbPath,
                })
                .Build();

            services.AddWpfBlazorWebView();
#if DEBUG
            services.AddBlazorWebViewDeveloperTools();
#endif

            services.AddLogging(builder =>
            {
                builder.SetMinimumLevel(LogLevel.Warning);
                builder.AddProvider(new StandaloneFileLoggerProvider());
            });

            // App services
            services.AddSingleton<IConfiguration>(configuration);
            services.AddSingleton(appPathsService);
            services.AddSingleton<AppActivityLogger>();
            services.AddSingleton<AiProviderSettingsService>();
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
            services.AddSingleton<BmesFcostActualService>();
            services.AddSingleton<BmesFCostReportCalculationService>();
            services.AddSingleton<BmesLpaScrapeService>();
            services.AddSingleton<BmesLpaImageService>();
            services.AddSingleton<BmesLpaHtmlExportService>();

            Services = services.BuildServiceProvider();

            base.OnStartup(e);
        }
        catch (Exception ex)
        {
            StandaloneErrorLog.Write("Startup failure", ex);
            MessageBox.Show(
                "BMES NG Rate could not start.\n\n" + ex.Message + "\n\nLog:\n" + StandaloneErrorLog.LogPath,
                "BMES NG Rate",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(-1);
        }
    }

    private void InstallGlobalExceptionHandlers()
    {
        DispatcherUnhandledException += (_, args) =>
        {
            StandaloneErrorLog.Write("DispatcherUnhandledException", args.Exception);
            MessageBox.Show(
                "An unexpected error occurred.\n\n" + args.Exception.Message + "\n\nLog:\n" + StandaloneErrorLog.LogPath,
                "BMES NG Rate",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            args.Handled = true;
        };

        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            if (args.ExceptionObject is Exception ex)
                StandaloneErrorLog.Write("UnhandledException", ex);
            else
                StandaloneErrorLog.Write("UnhandledException", args.ExceptionObject?.ToString() ?? "Unknown exception.");
        };

        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            StandaloneErrorLog.Write("UnobservedTaskException", args.Exception);
            args.SetObserved();
        };
    }
}
