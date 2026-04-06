using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Win32;
using QuickShareClone.Server;

namespace JinoSupporter.App.Modules.FileTransfer;

public partial class FileTransferView : UserControl
{
    public event Action? WebModuleSnapshotChanged;

    private readonly DispatcherTimer _refreshTimer;
    private readonly List<string> _selectedFiles = [];

    public FileTransferView()
    {
        InitializeComponent();
        _refreshTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _refreshTimer.Tick += async (_, _) => await RefreshAsync();
        Loaded += FileTransferView_Loaded;
        Unloaded += (_, _) => _refreshTimer.Stop();
    }

    public object GetWebModuleSnapshot()
    {
        return new
        {
            moduleType = "FileTransfer",
            deviceCount = DeviceListBox.Items.Count,
            selectedFileCount = _selectedFiles.Count,
            inboundSessionCount = InboundSessionListBox.Items.Count,
            transferCount = HistoryListBox.Items.Count,
            statusMessage = StatusTextBlock.Text
        };
    }

    public object UpdateWebModuleState(JsonElement payload) => GetWebModuleSnapshot();

    public object InvokeWebModuleAction(string action)
    {
        switch (action)
        {
            case "add-file":
                ChooseFilesButton_Click(this, new System.Windows.RoutedEventArgs());
                break;
            case "start-transfer":
                _ = SendSelectedFilesAsync();
                break;
        }

        return GetWebModuleSnapshot();
    }

    private async void FileTransferView_Loaded(object sender, System.Windows.RoutedEventArgs e)
    {
        await FileTransferRuntime.Instance.EnsureStartedAsync();
        _refreshTimer.Start();
        await RefreshAsync();
    }

    private async Task RefreshAsync()
    {
        try
        {
            await FileTransferRuntime.Instance.EnsureStartedAsync();
            DiscoveryDeviceInfo device = FileTransferRuntime.Instance.DeviceIdentityService.GetCurrentDevice();
            LocalDeviceTextBlock.Text = device.DeviceName;
            LocalUrlTextBlock.Text = string.Join(Environment.NewLine, device.ServerUrls);

            DeviceListBox.ItemsSource = FileTransferRuntime.Instance.AndroidDeviceStore.List()
                .Select(item => new DisplayItem(item.DeviceId, $"{item.DeviceName}{Environment.NewLine}{item.ReceiveUrl}"))
                .ToArray();

            InboundSessionListBox.ItemsSource = FileTransferRuntime.Instance.UploadStore.List()
                .Select(item => new DisplayItem(item.FileId, BuildInboundText(item)))
                .ToArray();

            HistoryListBox.ItemsSource = FileTransferRuntime.Instance.AndroidOutboundTransferStore.List()
                .Select(item => new DisplayItem(item.TransferId, BuildOutboundText(item)))
                .ToArray();

            SelectedFileListBox.ItemsSource = _selectedFiles.Select(path => new DisplayItem(path, path)).ToArray();
            SetStatus("File transfer runtime ready.");
        }
        catch (Exception ex)
        {
            SetStatus(ex.Message);
        }
    }

    private void RefreshButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        _ = RefreshAsync();
    }

    private void ChooseFilesButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Multiselect = true,
            Title = "Choose files to send to Android"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        _selectedFiles.Clear();
        _selectedFiles.AddRange(dialog.FileNames);
        SelectedFileListBox.ItemsSource = _selectedFiles.Select(path => new DisplayItem(path, path)).ToArray();
        SetStatus($"{_selectedFiles.Count} file(s) selected.");
    }

    private async void SendFilesButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        await SendSelectedFilesAsync();
    }

    private async Task SendSelectedFilesAsync()
    {
        if (DeviceListBox.SelectedItem is not DisplayItem selectedDevice)
        {
            SetStatus("Choose an Android device first.");
            return;
        }

        if (_selectedFiles.Count == 0)
        {
            SetStatus("Choose at least one file first.");
            return;
        }

        try
        {
            IsEnabled = false;
            SetStatus("Sending files to Android...");
            await FileTransferRuntime.Instance.SendFilesToAndroidAsync(selectedDevice.Id, _selectedFiles, CancellationToken.None);
            await RefreshAsync();
            SetStatus("File transfer request finished.");
        }
        catch (Exception ex)
        {
            SetStatus(ex.Message);
        }
        finally
        {
            IsEnabled = true;
        }
    }

    private void ChooseDestinationButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (InboundSessionListBox.SelectedItem is not DisplayItem selectedSession)
        {
            SetStatus("Choose an inbound session first.");
            return;
        }

        using System.Windows.Forms.FolderBrowserDialog dialog = new()
        {
            Description = "Choose a destination folder for received files",
            ShowNewFolderButton = true
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
        {
            return;
        }

        FileTransferRuntime.Instance.UploadStore.SetDestination(selectedSession.Id, dialog.SelectedPath);
        _ = RefreshAsync();
        SetStatus("Destination folder saved.");
    }

    private void SetStatus(string message)
    {
        StatusTextBlock.Text = message;
        NotifyWebModuleSnapshotChanged();
    }

    private void NotifyWebModuleSnapshotChanged()
    {
        WebModuleSnapshotChanged?.Invoke();
    }

    private static string BuildInboundText(UploadSessionSummary item)
    {
        string state = item.IsCompleted ? "Completed" : item.DestinationSelected ? "Receiving" : "Waiting for folder";
        return $"{item.FileName}{Environment.NewLine}{state} | {item.ReceivedBytes} / {item.TotalBytes.GetValueOrDefault()} bytes";
    }

    private static string BuildOutboundText(AndroidOutboundTransferSummary item)
    {
        return $"{item.FileName}{Environment.NewLine}{item.DeviceName} | {item.StatusText} | {item.SentBytes} / {item.TotalBytes} bytes";
    }

    private sealed record DisplayItem(string Id, string Text)
    {
        public override string ToString() => Text;
    }
}
