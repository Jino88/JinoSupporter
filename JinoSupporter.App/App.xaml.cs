using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using DataMaker.Logger;
using JinoSupporter.App.Modules.FileTransfer;

namespace JinoSupporter.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnCurrentDomainUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        base.OnStartup(e);
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        try
        {
            if (FileTransferRuntime.IsCreated)
            {
                await FileTransferRuntime.Instance.DisposeAsync();
            }
        }
        catch (Exception ex)
        {
            clLogger.LogException(ex, "FileTransfer runtime shutdown");
        }
        finally
        {
            base.OnExit(e);
        }
    }

    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        clLogger.LogException(e.Exception, "Global UI exception");
        MessageBox.Show(
            $"Unexpected error was caught and the app will continue.\n\n{e.Exception.Message}",
            "Application Warning",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
        e.Handled = true;
    }

    private static void OnCurrentDomainUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            clLogger.LogException(ex, "AppDomain unhandled exception");
        }
    }

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        clLogger.LogException(e.Exception, "Unobserved task exception");
        e.SetObserved();
    }
}
