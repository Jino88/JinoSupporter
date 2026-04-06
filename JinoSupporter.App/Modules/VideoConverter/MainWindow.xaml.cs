using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using WorkbenchHost.Infrastructure;

namespace VideoConverter;

public partial class MainWindow : Window
{
    private static readonly string[] SupportedExtensions = [".mp4", ".mkv", ".avi", ".mov", ".ts", ".flv", ".wmv", ".m4v"];
    public static string AppDataDirectory => AppSettingsPathManager.GetModuleDirectory("VideoConverter");
    public static string SettingsFilePath => WorkbenchSettingsStore.SettingsFilePath;
    private readonly ObservableCollection<EncodingQueueItem> _items = [];

    private CancellationTokenSource? _encodingCts;
    private CancellationTokenSource? _thumbnailCts;
    private CancellationTokenSource? _detailCts;
    private string? _thumbnailTempPath;
    private bool _isEncoding;
    private bool _isSyncingCrf;
    private bool _isInitializing = true;
    private int _completedCount;
    private int _failedCount;

    public event Action? WebModuleSnapshotChanged;

    public MainWindow()
    {
        InitializeComponent();
        QueueDataGrid.ItemsSource = _items;
        _items.CollectionChanged += Items_CollectionChanged;

        CodecComboBox.ItemsSource = new[]
        {
            new CodecOption("H.264 (CPU, libx264)", "libx264"),
            new CodecOption("HEVC/H.265 (CPU, libx265)", "libx265")
        };
        CodecComboBox.SelectedIndex = 0;

        ResolutionComboBox.ItemsSource = new[]
        {
            new ResolutionOption("Keep Source", string.Empty),
            new ResolutionOption("1920x1080", "scale=1920:1080:flags=lanczos"),
            new ResolutionOption("1280x720", "scale=1280:720:flags=lanczos"),
            new ResolutionOption("854x480", "scale=854:480:flags=lanczos")
        };
        ResolutionComboBox.SelectedIndex = 0;

        PresetComboBox.ItemsSource = new[]
        {
            new PresetOption("medium", "medium"),
            new PresetOption("fast", "fast"),
            new PresetOption("slow", "slow"),
            new PresetOption("veryfast", "veryfast")
        };
        PresetComboBox.SelectedIndex = 0;

        AudioBitrateComboBox.ItemsSource = new[] { "96k", "128k", "192k", "256k", "320k" };
        AudioBitrateComboBox.SelectedItem = "192k";

        FfmpegFolderTextBox.Text = Environment.CurrentDirectory;
        OutputFolderTextBox.Text = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            "EncodedVideos");

        _isSyncingCrf = true;
        CrfSlider.Value = 21;
        CrfTextBox.Text = "21";
        _isSyncingCrf = false;

        UpdateStateText();
        ClearSelectedVideoDetails();
        LoadSettings();
        _isInitializing = false;
        NotifyWebModuleSnapshotChanged();
    }

    private void Items_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems is not null)
        {
            foreach (EncodingQueueItem item in e.OldItems.OfType<EncodingQueueItem>())
            {
                item.PropertyChanged -= QueueItem_PropertyChanged;
            }
        }

        if (e.NewItems is not null)
        {
            foreach (EncodingQueueItem item in e.NewItems.OfType<EncodingQueueItem>())
            {
                item.PropertyChanged += QueueItem_PropertyChanged;
            }
        }

        NotifyWebModuleSnapshotChanged();
    }

    private void QueueItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        NotifyWebModuleSnapshotChanged();
    }

    private void AddFilesButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Video Files|*.mp4;*.mkv;*.avi;*.mov;*.ts;*.flv;*.wmv;*.m4v|All Files|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog(GetDialogOwner()) == true)
        {
            _ = AddFilesAsync(dialog.FileNames);
        }
    }

    private void AddFolderButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select video source folder",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = false
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        IEnumerable<string> files = Directory
            .EnumerateFiles(dialog.SelectedPath, "*.*", SearchOption.AllDirectories)
            .Where(path => SupportedExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase));
        _ = AddFilesAsync(files);
    }

    private void RemoveSelectedButton_Click(object sender, RoutedEventArgs e)
    {
        List<EncodingQueueItem> selectedItems = QueueDataGrid.SelectedItems.Cast<EncodingQueueItem>().ToList();
        foreach (EncodingQueueItem item in selectedItems)
        {
            _items.Remove(item);
        }

        if (_items.Count == 0)
        {
            ClearThumbnailPreview();
            ClearSelectedVideoDetails();
        }
        else if (QueueDataGrid.SelectedItem is null)
        {
            QueueDataGrid.SelectedIndex = 0;
        }

        UpdateStateText();
        NotifyWebModuleSnapshotChanged();
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isEncoding)
        {
            return;
        }

        _items.Clear();
        _completedCount = 0;
        _failedCount = 0;
        ClearThumbnailPreview();
        ClearSelectedVideoDetails();
        UpdateStateText();
        NotifyWebModuleSnapshotChanged();
    }

    private async void StartButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isEncoding)
        {
            return;
        }

        if (_items.Count == 0)
        {
            MessageBox.Show(GetDialogOwner(), "Please add video files first.", "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (!TryGetSettings(out EncodingSettings settings))
        {
            return;
        }

        Directory.CreateDirectory(settings.OutputFolder);

        _completedCount = 0;
        _failedCount = 0;
        _isEncoding = true;
        _encodingCts = new CancellationTokenSource();
        StartButton.IsEnabled = false;
        CancelButton.IsEnabled = true;

        foreach (EncodingQueueItem item in _items)
        {
            item.Progress = 0;
            item.Status = "Queued";
            item.OutputPath = string.Empty;
        }

        UpdateStateText();

        try
        {
            foreach (EncodingQueueItem item in _items)
            {
                if (settings.HevcOnlyMode && !item.IsHevc)
                {
                    item.Status = "Skipped (not HEVC)";
                    item.Progress = 0;
                    continue;
                }

                _encodingCts.Token.ThrowIfCancellationRequested();
                await EncodeOneAsync(item, settings, _encodingCts.Token);
            }
        }
        catch (OperationCanceledException)
        {
            foreach (EncodingQueueItem item in _items.Where(i => i.Status == "Running"))
            {
                item.Status = "Canceled";
            }
        }
        finally
        {
            _encodingCts.Dispose();
            _encodingCts = null;
            _isEncoding = false;
            StartButton.IsEnabled = true;
            CancelButton.IsEnabled = false;
            UpdateStateText();
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _encodingCts?.Cancel();
    }

    private void BrowseFfmpegFolderButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select folder containing ffmpeg and ffprobe",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = false
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            FfmpegFolderTextBox.Text = dialog.SelectedPath;
            _ = RefreshSelectedThumbnailPreviewAsync();
            _ = RefreshSelectedVideoDetailsAsync();
        }
    }

    private void BrowseOutputFolderButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select output folder",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            OutputFolderTextBox.Text = dialog.SelectedPath;
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void CrfSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (_isSyncingCrf)
        {
            return;
        }

        _isSyncingCrf = true;
        CrfTextBox.Text = ((int)e.NewValue).ToString(CultureInfo.InvariantCulture);
        _isSyncingCrf = false;
        NotifyWebModuleSnapshotChanged();
    }

    private void CrfTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (_isSyncingCrf)
        {
            return;
        }

        if (!int.TryParse(CrfTextBox.Text.Trim(), out int crf))
        {
            return;
        }

        crf = Math.Clamp(crf, 0, 51);
        _isSyncingCrf = true;
        CrfSlider.Value = crf;
        if (!string.Equals(CrfTextBox.Text, crf.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))
        {
            CrfTextBox.Text = crf.ToString(CultureInfo.InvariantCulture);
            CrfTextBox.SelectionStart = CrfTextBox.Text.Length;
        }
        _isSyncingCrf = false;
        NotifyWebModuleSnapshotChanged();
    }

    private async void QueueDataGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        await Task.WhenAll(RefreshSelectedThumbnailPreviewAsync(), RefreshSelectedVideoDetailsAsync());
        NotifyWebModuleSnapshotChanged();
    }

    private async void CropSetting_Changed(object sender, RoutedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }

        await RefreshSelectedThumbnailPreviewAsync();
        NotifyWebModuleSnapshotChanged();
    }

    private async void FfmpegFolderTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (_isInitializing)
        {
            return;
        }

        SaveSettings();
        _ = ProbeQueueItemsAsync(_items);
        await Task.WhenAll(RefreshSelectedThumbnailPreviewAsync(), RefreshSelectedVideoDetailsAsync());
        NotifyWebModuleSnapshotChanged();
    }

    private void QueueDataGrid_DragOver(object sender, DragEventArgs e)
    {
        bool hasFileDrop = e.Data.GetDataPresent(DataFormats.FileDrop);
        e.Effects = hasFileDrop ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void QueueDataGrid_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            return;
        }

        if (e.Data.GetData(DataFormats.FileDrop) is not string[] droppedPaths || droppedPaths.Length == 0)
        {
            return;
        }

        var files = new List<string>();
        foreach (string path in droppedPaths)
        {
            if (File.Exists(path))
            {
                string ext = Path.GetExtension(path);
                if (SupportedExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
                {
                    files.Add(path);
                }
                continue;
            }

            if (Directory.Exists(path))
            {
                IEnumerable<string> folderFiles = Directory
                    .EnumerateFiles(path, "*.*", SearchOption.AllDirectories)
                    .Where(file => SupportedExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase));
                files.AddRange(folderFiles);
            }
        }

        _ = AddFilesAsync(files);
    }

    private async Task RefreshSelectedThumbnailPreviewAsync()
    {
        string? selectedPath = QueueDataGrid.SelectedItem is EncodingQueueItem item ? item.InputPath : null;
        await UpdateThumbnailPreviewAsync(selectedPath);
    }

    private async Task RefreshSelectedVideoDetailsAsync()
    {
        string? selectedPath = QueueDataGrid.SelectedItem is EncodingQueueItem item ? item.InputPath : null;
        await UpdateSelectedVideoDetailsAsync(selectedPath);
    }

    private void SaveSettings()
    {
        try
        {
            WorkbenchSettingsStore.UpdateSettings(store =>
            {
                store.VideoConverter.FfmpegFolder = FfmpegFolderTextBox.Text.Trim();
                store.VideoConverter.OutputFolder = OutputFolderTextBox.Text.Trim();
                store.VideoConverter.HevcOnlyMode = HevcOnlyModeCheckBox.IsChecked == true;
            });
        }
        catch
        {
        }
    }

    private void LoadSettings()
    {
        try
        {
            WorkbenchSettingsStore.VideoConverterSetting settings = WorkbenchSettingsStore.GetSettings().VideoConverter;

            if (!string.IsNullOrWhiteSpace(settings.FfmpegFolder))
            {
                FfmpegFolderTextBox.Text = settings.FfmpegFolder;
            }

            if (!string.IsNullOrWhiteSpace(settings.OutputFolder))
            {
                OutputFolderTextBox.Text = settings.OutputFolder;
            }

            HevcOnlyModeCheckBox.IsChecked = settings.HevcOnlyMode;
        }
        catch
        {
        }
    }

    private async Task ProbeQueueItemsAsync(IEnumerable<EncodingQueueItem> items)
    {
        if (!TryResolveFfmpegTools(FfmpegFolderTextBox.Text.Trim(), out _, out string ffprobeExe, out _))
        {
            return;
        }

        foreach (EncodingQueueItem item in items)
        {
            if (!File.Exists(item.InputPath))
            {
                continue;
            }

            VideoProbeInfo? info = await ProbeVideoInfoAsync(ffprobeExe, item.InputPath, CancellationToken.None);
            item.VideoCodec = info?.VideoCodec ?? "-";
            item.IsHevc = IsHevcCodec(info?.VideoCodec);

            if (item.Status == "Queued" || item.Status == "Ready" || item.Status == "Ready (HEVC)")
            {
                item.Status = item.IsHevc ? "Ready (HEVC)" : "Ready";
            }
        }
    }

    private async Task AddFilesAsync(IEnumerable<string> filePaths)
    {
        var addedItems = new List<EncodingQueueItem>();

        foreach (string path in filePaths)
        {
            if (!File.Exists(path))
            {
                continue;
            }

            if (_items.Any(item => string.Equals(item.InputPath, path, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var item = new EncodingQueueItem(path);
            _items.Add(item);
            addedItems.Add(item);
        }

        if (_items.Count > 0 && QueueDataGrid.SelectedItem is null)
        {
            QueueDataGrid.SelectedIndex = 0;
        }

        UpdateStateText();

        if (addedItems.Count > 0)
        {
            await ProbeQueueItemsAsync(addedItems);
        }

        NotifyWebModuleSnapshotChanged();
    }

    private bool TryGetSettings(out EncodingSettings settings)
    {
        settings = default;

        string ffmpegFolder = FfmpegFolderTextBox.Text.Trim();
        string outputFolder = OutputFolderTextBox.Text.Trim();
        SaveSettings();

        if (!TryResolveFfmpegTools(ffmpegFolder, out string ffmpegPath, out string ffprobePath, out string errorMessage))
        {
            MessageBox.Show(GetDialogOwner(), errorMessage, "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            MessageBox.Show(GetDialogOwner(), "Please enter output folder.", "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (!int.TryParse(CrfTextBox.Text.Trim(), out int crf) || crf is < 0 or > 51)
        {
            MessageBox.Show(GetDialogOwner(), "CRF must be integer between 0 and 51.", "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (CodecComboBox.SelectedItem is not CodecOption codec)
        {
            MessageBox.Show(GetDialogOwner(), "Select a codec.", "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (ResolutionComboBox.SelectedItem is not ResolutionOption resolution)
        {
            MessageBox.Show(GetDialogOwner(), "Select a resolution option.", "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (PresetComboBox.SelectedItem is not PresetOption preset)
        {
            MessageBox.Show(GetDialogOwner(), "Select a preset.", "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        int cropX = 0;
        int cropY = 0;
        int cropW = 0;
        int cropH = 0;
        bool enableCrop = EnableCropCheckBox.IsChecked == true;

        if (enableCrop)
        {
            bool parsed = int.TryParse(CropXTextBox.Text.Trim(), out cropX)
                && int.TryParse(CropYTextBox.Text.Trim(), out cropY)
                && int.TryParse(CropWTextBox.Text.Trim(), out cropW)
                && int.TryParse(CropHTextBox.Text.Trim(), out cropH);

            if (!parsed || cropW <= 0 || cropH <= 0 || cropX < 0 || cropY < 0)
            {
                MessageBox.Show(GetDialogOwner(), "Crop values must be integers and W/H must be > 0.", "VideoConverter", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
        }

        settings = new EncodingSettings(
            ffmpegPath,
            ffprobePath,
            outputFolder,
            codec.Codec,
            resolution.FilterExpression,
            preset.Preset,
            crf,
            enableCrop,
            cropX,
            cropY,
            cropW,
            cropH,
            DisableAudioCheckBox.IsChecked == true,
            (AudioBitrateComboBox.SelectedItem as string) ?? "192k",
            OverwriteCheckBox.IsChecked == true,
            HevcOnlyModeCheckBox.IsChecked == true);

        return true;
    }

    private async Task EncodeOneAsync(EncodingQueueItem item, EncodingSettings settings, CancellationToken ct)
    {
        item.Status = "Running";
        UpdateStateText();

        string outputFileName = $"{Path.GetFileNameWithoutExtension(item.InputPath)}_encoded.mp4";
        string outputPath = Path.Combine(settings.OutputFolder, outputFileName);
        outputPath = EnsureUniquePath(outputPath, settings.Overwrite);
        item.OutputPath = outputPath;

        double totalDuration = await ProbeDurationAsync(settings.FfprobePath, item.InputPath, ct);
        string overwriteArg = settings.Overwrite ? "-y" : "-n";

        var filters = new List<string>();
        if (settings.EnableCrop)
        {
            filters.Add($"crop={settings.CropW}:{settings.CropH}:{settings.CropX}:{settings.CropY}");
        }

        if (!string.IsNullOrWhiteSpace(settings.ResolutionFilterExpression))
        {
            filters.Add(settings.ResolutionFilterExpression);
        }

        string videoFilterArg = filters.Count > 0
            ? $"-vf \"{string.Join(",", filters)}\" "
            : string.Empty;

        string audioArg = settings.DisableAudio
            ? "-an"
            : $"-c:a aac -b:a {settings.AudioBitrate}";

        string arguments =
            $"{overwriteArg} -hide_banner -i \"{item.InputPath}\" -c:v {settings.VideoCodec} -preset {settings.Preset} -crf {settings.Crf} " +
            $"{videoFilterArg}{audioArg} -movflags +faststart " +
            $"\"{outputPath}\"";

        var startInfo = new ProcessStartInfo
        {
            FileName = settings.FfmpegPath,
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };

        using var process = new Process { StartInfo = startInfo };
        process.Start();

        using var registration = ct.Register(() =>
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
            }
        });

        while (!process.HasExited)
        {
            string? line = await process.StandardError.ReadLineAsync();
            if (line is null)
            {
                break;
            }

            if (totalDuration > 0)
            {
                Match match = Regex.Match(line, @"time=(\d+):(\d+):([\d\.]+)");
                if (match.Success)
                {
                    double hour = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                    double minute = double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
                    double second = double.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture);
                    double current = hour * 3600 + minute * 60 + second;
                    item.Progress = Math.Clamp(current / totalDuration * 100.0, 0, 100);
                }
            }
        }

        string stderr = await process.StandardError.ReadToEndAsync();
        _ = await process.StandardOutput.ReadToEndAsync();
        await process.WaitForExitAsync(ct);

        if (ct.IsCancellationRequested)
        {
            item.Status = "Canceled";
            UpdateStateText();
            return;
        }

        if (process.ExitCode == 0 && File.Exists(outputPath))
        {
            item.Progress = 100;
            item.Status = "Completed";
            _completedCount++;
        }
        else
        {
            item.Status = "Failed";
            _failedCount++;
            if (string.IsNullOrWhiteSpace(stderr))
            {
                item.ErrorMessage = "Unknown FFmpeg error";
            }
            else
            {
                item.ErrorMessage = stderr.Length > 200 ? stderr[..200] : stderr;
            }
        }

        UpdateStateText();
    }

    private static string EnsureUniquePath(string outputPath, bool overwrite)
    {
        if (overwrite || !File.Exists(outputPath))
        {
            return outputPath;
        }

        string directory = Path.GetDirectoryName(outputPath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(outputPath);
        string extension = Path.GetExtension(outputPath);

        int suffix = 1;
        string candidate = outputPath;
        while (File.Exists(candidate))
        {
            candidate = Path.Combine(directory, $"{fileName}_{suffix}{extension}");
            suffix++;
        }

        return candidate;
    }

    private async Task<double> ProbeDurationAsync(string ffprobePath, string inputPath, CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = ffprobePath,
            Arguments = $"-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 \"{inputPath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        try
        {
            using var proc = new Process { StartInfo = psi };
            proc.Start();
            string output = (await proc.StandardOutput.ReadToEndAsync()).Trim();
            _ = await proc.StandardError.ReadToEndAsync();
            await proc.WaitForExitAsync(ct);

            if (double.TryParse(output, NumberStyles.Any, CultureInfo.InvariantCulture, out double duration))
            {
                return duration;
            }
        }
        catch
        {
        }

        return 0;
    }

    private async Task UpdateThumbnailPreviewAsync(string? inputPath)
    {
        _thumbnailCts?.Cancel();
        _thumbnailCts?.Dispose();
        _thumbnailCts = new CancellationTokenSource();
        CancellationToken ct = _thumbnailCts.Token;

        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            ClearThumbnailPreview();
            return;
        }

        ThumbnailInfoText.Text = $"Loading thumbnail: {Path.GetFileName(inputPath)}";

        string ffmpegFolder = FfmpegFolderTextBox.Text.Trim();
        if (!TryResolveFfmpegTools(ffmpegFolder, out string ffmpegExe, out _, out string errorMessage))
        {
            ThumbnailInfoText.Text = errorMessage;
            return;
        }

        string previewFilter = BuildThumbnailFilterDescription(out string cropDescription);
        string tempPath = Path.Combine(Path.GetTempPath(), $"vh_preview_{Guid.NewGuid():N}.jpg");
        bool created = await ExtractThumbnailAsync(ffmpegExe, inputPath, tempPath, previewFilter, ct);
        if (!created)
        {
            created = await ExtractThumbnailAsync(ffmpegExe, inputPath, tempPath, previewFilter, ct, seekSeconds: 0.0);
        }

        if (!created || ct.IsCancellationRequested || !File.Exists(tempPath))
        {
            ThumbnailInfoText.Text = $"Thumbnail unavailable: {Path.GetFileName(inputPath)}";
            TryDeleteFile(tempPath);
            return;
        }

        try
        {
            using var stream = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.StreamSource = stream;
            bitmap.EndInit();
            bitmap.Freeze();

            ThumbnailPreviewImage.Source = bitmap;
            ThumbnailInfoText.Text = $"Preview: {Path.GetFileName(inputPath)}{cropDescription}";
        }
        catch
        {
            ThumbnailInfoText.Text = $"Thumbnail load failed: {Path.GetFileName(inputPath)}";
            TryDeleteFile(tempPath);
            return;
        }

        if (!string.IsNullOrWhiteSpace(_thumbnailTempPath))
        {
            TryDeleteFile(_thumbnailTempPath);
        }
        _thumbnailTempPath = tempPath;
    }

    private string BuildThumbnailFilterDescription(out string cropDescription)
    {
        cropDescription = string.Empty;

        bool enableCrop = EnableCropCheckBox.IsChecked == true;
        bool validCrop = enableCrop
            && int.TryParse(CropXTextBox.Text.Trim(), out int x)
            && int.TryParse(CropYTextBox.Text.Trim(), out int y)
            && int.TryParse(CropWTextBox.Text.Trim(), out int w)
            && int.TryParse(CropHTextBox.Text.Trim(), out int h)
            && x >= 0
            && y >= 0
            && w > 0
            && h > 0;

        if (!enableCrop)
        {
            return "scale=320:-2";
        }

        if (!validCrop)
        {
            cropDescription = " (Crop values invalid)";
            return "scale=320:-2";
        }

        int cropX = int.Parse(CropXTextBox.Text.Trim(), CultureInfo.InvariantCulture);
        int cropY = int.Parse(CropYTextBox.Text.Trim(), CultureInfo.InvariantCulture);
        int cropW = int.Parse(CropWTextBox.Text.Trim(), CultureInfo.InvariantCulture);
        int cropH = int.Parse(CropHTextBox.Text.Trim(), CultureInfo.InvariantCulture);

        cropDescription = $" (Crop: X={cropX}, Y={cropY}, W={cropW}, H={cropH})";
        return $"drawbox=x={cropX}:y={cropY}:w={cropW}:h={cropH}:color=yellow@0.9:t=3,scale=320:-2";
    }

    private static async Task<bool> ExtractThumbnailAsync(string ffmpegExe, string inputPath, string outputImagePath, string videoFilter, CancellationToken ct, double seekSeconds = 1.0)
    {
        string args = $"-y -hide_banner -loglevel error -ss {seekSeconds.ToString("F3", CultureInfo.InvariantCulture)} -i \"{inputPath}\" -frames:v 1 -vf \"{videoFilter}\" \"{outputImagePath}\"";
        var psi = new ProcessStartInfo
        {
            FileName = ffmpegExe,
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        try
        {
            using var proc = new Process { StartInfo = psi };
            proc.Start();
            _ = await proc.StandardOutput.ReadToEndAsync();
            _ = await proc.StandardError.ReadToEndAsync();
            await proc.WaitForExitAsync(ct);
            return proc.ExitCode == 0 && File.Exists(outputImagePath);
        }
        catch
        {
            return false;
        }
    }

    private async Task UpdateSelectedVideoDetailsAsync(string? inputPath)
    {
        _detailCts?.Cancel();
        _detailCts?.Dispose();
        _detailCts = new CancellationTokenSource();
        CancellationToken ct = _detailCts.Token;

        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            ClearSelectedVideoDetails();
            return;
        }

        DetailFileNameTextBlock.Text = Path.GetFileName(inputPath);
        DetailPathTextBlock.Text = inputPath;
        DetailSummaryTextBlock.Text = "Loading media information...";
        DetailVideoTextBlock.Text = string.Empty;
        DetailAudioTextBlock.Text = string.Empty;

        if (!TryResolveFfmpegTools(FfmpegFolderTextBox.Text.Trim(), out _, out string ffprobeExe, out string errorMessage))
        {
            DetailSummaryTextBlock.Text = errorMessage;
            return;
        }

        VideoProbeInfo? info = await ProbeVideoInfoAsync(ffprobeExe, inputPath, ct);
        if (ct.IsCancellationRequested)
        {
            return;
        }

        if (info is not VideoProbeInfo probeInfo)
        {
            DetailSummaryTextBlock.Text = "Cannot read metadata from selected video.";
            return;
        }

        DetailSummaryTextBlock.Text =
            $"Duration: {FormatDuration(probeInfo.DurationSeconds)} | File Size: {FormatBytes(probeInfo.FileSizeBytes)} | Container Bitrate: {FormatBitrate(probeInfo.FormatBitrate)}";
        DetailVideoTextBlock.Text =
            $"Video: {probeInfo.VideoCodec ?? "-"} | HEVC: {(IsHevcCodec(probeInfo.VideoCodec) ? "Yes" : "No")} | Resolution: {(probeInfo.Width > 0 && probeInfo.Height > 0 ? $"{probeInfo.Width}x{probeInfo.Height}" : "-")} | FPS: {FormatFps(probeInfo.Fps)} | Bitrate: {FormatBitrate(probeInfo.VideoBitrate)}";
        DetailAudioTextBlock.Text =
            $"Audio: {probeInfo.AudioCodec ?? "-"} | Sample Rate: {(probeInfo.AudioSampleRate > 0 ? $"{probeInfo.AudioSampleRate} Hz" : "-")} | Channels: {(probeInfo.AudioChannels > 0 ? probeInfo.AudioChannels.ToString(CultureInfo.InvariantCulture) : "-")} | Bitrate: {FormatBitrate(probeInfo.AudioBitrate)}";
    }

    private void ClearThumbnailPreview()
    {
        if (ThumbnailPreviewImage is not null)
        {
            ThumbnailPreviewImage.Source = null;
        }

        if (ThumbnailInfoText is not null)
        {
            ThumbnailInfoText.Text = "Select a video from the left list to preview thumbnail.";
        }

        if (!string.IsNullOrWhiteSpace(_thumbnailTempPath))
        {
            TryDeleteFile(_thumbnailTempPath);
            _thumbnailTempPath = null;
        }

        NotifyWebModuleSnapshotChanged();
    }

    private void ClearSelectedVideoDetails()
    {
        DetailFileNameTextBlock.Text = "-";
        DetailPathTextBlock.Text = "-";
        DetailSummaryTextBlock.Text = "Select a video to see details.";
        DetailVideoTextBlock.Text = string.Empty;
        DetailAudioTextBlock.Text = string.Empty;
        NotifyWebModuleSnapshotChanged();
    }

    private static void TryDeleteFile(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return;
        }

        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch
        {
        }
    }

    private void UpdateStateText()
    {
        int queued = _items.Count(item => item.Status == "Queued");
        int running = _items.Count(item => item.Status == "Running");
        int canceled = _items.Count(item => item.Status == "Canceled");
        bool hevcOnlyMode = HevcOnlyModeCheckBox.IsChecked == true;
        int targetCount = hevcOnlyMode
            ? _items.Count(item => item.IsHevc)
            : _items.Count;
        int doneCount = _items.Count(item =>
            item.Status == "Completed"
            || item.Status == "Failed"
            || item.Status == "Canceled");
        if (hevcOnlyMode)
        {
            doneCount += _items.Count(item => item.Status == "Skipped (not HEVC)");
        }

        double percent = targetCount > 0
            ? Math.Clamp((_completedCount + _failedCount + canceled) / (double)targetCount * 100.0, 0, 100)
            : 0;

        string progressText = hevcOnlyMode
            ? $"HEVC Progress: {Math.Min(_completedCount + _failedCount + canceled, targetCount)}/{targetCount} ({percent:F1}%)"
            : $"Progress: {Math.Min(doneCount, targetCount)}/{targetCount}";

        StateTextBlock.Text =
            $"{progressText} | Queued: {queued}, Running: {running}, Completed: {_completedCount}, Failed: {_failedCount}, Canceled: {canceled}, Total: {_items.Count}";
        NotifyWebModuleSnapshotChanged();
    }

    private Window? GetDialogOwner()
    {
        Window? hostWindow = Application.Current?.MainWindow;
        if (hostWindow is not null && hostWindow.IsVisible)
        {
            return hostWindow;
        }

        return null;
    }

    protected override void OnClosed(EventArgs e)
    {
        _items.CollectionChanged -= Items_CollectionChanged;
        foreach (EncodingQueueItem item in _items)
        {
            item.PropertyChanged -= QueueItem_PropertyChanged;
        }

        _encodingCts?.Cancel();
        _thumbnailCts?.Cancel();
        _detailCts?.Cancel();
        _thumbnailCts?.Dispose();
        _detailCts?.Dispose();
        _encodingCts?.Dispose();
        ClearThumbnailPreview();
        base.OnClosed(e);
    }

    public object GetWebModuleSnapshot()
    {
        EncodingQueueItem? selectedItem = QueueDataGrid.SelectedItem as EncodingQueueItem;
        BitmapSource? previewSource = ThumbnailPreviewImage.Source as BitmapSource;

        return new
        {
            moduleType = "VideoConverter",
            stateText = StateTextBlock.Text ?? string.Empty,
            isEncoding = _isEncoding,
            canStart = StartButton.IsEnabled,
            canCancel = CancelButton.IsEnabled,
            ffmpegFolder = FfmpegFolderTextBox.Text ?? string.Empty,
            outputFolder = OutputFolderTextBox.Text ?? string.Empty,
            codec = (CodecComboBox.SelectedItem as CodecOption)?.Codec ?? string.Empty,
            codecLabel = (CodecComboBox.SelectedItem as CodecOption)?.DisplayName ?? string.Empty,
            resolution = (ResolutionComboBox.SelectedItem as ResolutionOption)?.FilterExpression ?? string.Empty,
            resolutionLabel = (ResolutionComboBox.SelectedItem as ResolutionOption)?.DisplayName ?? string.Empty,
            preset = (PresetComboBox.SelectedItem as PresetOption)?.Preset ?? string.Empty,
            presetLabel = (PresetComboBox.SelectedItem as PresetOption)?.DisplayName ?? string.Empty,
            crf = (int)Math.Round(CrfSlider.Value),
            enableCrop = EnableCropCheckBox.IsChecked == true,
            cropX = CropXTextBox.Text ?? "0",
            cropY = CropYTextBox.Text ?? "0",
            cropW = CropWTextBox.Text ?? "0",
            cropH = CropHTextBox.Text ?? "0",
            disableAudio = DisableAudioCheckBox.IsChecked == true,
            overwrite = OverwriteCheckBox.IsChecked == true,
            hevcOnlyMode = HevcOnlyModeCheckBox.IsChecked == true,
            audioBitrate = AudioBitrateComboBox.SelectedItem as string ?? "192k",
            selectedPath = selectedItem?.InputPath ?? string.Empty,
            thumbnailInfo = ThumbnailInfoText.Text ?? string.Empty,
            thumbnailDataUrl = previewSource is null ? string.Empty : ToDataUrl(previewSource),
            detailFileName = DetailFileNameTextBlock.Text ?? string.Empty,
            detailPath = DetailPathTextBlock.Text ?? string.Empty,
            detailSummary = DetailSummaryTextBlock.Text ?? string.Empty,
            detailVideo = DetailVideoTextBlock.Text ?? string.Empty,
            detailAudio = DetailAudioTextBlock.Text ?? string.Empty,
            queueItems = _items.Take(120).Select(item => new
            {
                fileName = item.FileName,
                inputPath = item.InputPath,
                videoCodec = item.VideoCodec,
                isHevcDisplay = item.IsHevcDisplay,
                status = item.Status,
                progress = item.Progress,
                progressText = item.ProgressText
            }).ToArray(),
            codecOptions = CodecComboBox.Items.OfType<CodecOption>().Select(option => new { label = option.DisplayName, value = option.Codec }).ToArray(),
            resolutionOptions = ResolutionComboBox.Items.OfType<ResolutionOption>().Select(option => new { label = option.DisplayName, value = option.FilterExpression }).ToArray(),
            presetOptions = PresetComboBox.Items.OfType<PresetOption>().Select(option => new { label = option.DisplayName, value = option.Preset }).ToArray(),
            bitrateOptions = AudioBitrateComboBox.Items.OfType<string>().ToArray()
        };
    }

    public object UpdateWebModuleState(JsonElement payload)
    {
        if (payload.TryGetProperty("selectedPath", out JsonElement selectedPathElement))
        {
            string? selectedPath = selectedPathElement.GetString();
            EncodingQueueItem? selectedItem = _items.FirstOrDefault(item =>
                string.Equals(item.InputPath, selectedPath, StringComparison.OrdinalIgnoreCase));
            if (selectedItem is not null)
            {
                QueueDataGrid.SelectedItem = selectedItem;
                QueueDataGrid.ScrollIntoView(selectedItem);
            }
        }

        if (payload.TryGetProperty("ffmpegFolder", out JsonElement ffmpegFolderElement))
        {
            FfmpegFolderTextBox.Text = ffmpegFolderElement.GetString() ?? string.Empty;
        }

        if (payload.TryGetProperty("outputFolder", out JsonElement outputFolderElement))
        {
            OutputFolderTextBox.Text = outputFolderElement.GetString() ?? string.Empty;
        }

        if (payload.TryGetProperty("codec", out JsonElement codecElement))
        {
            string? codec = codecElement.GetString();
            CodecOption? option = CodecComboBox.Items.OfType<CodecOption>().FirstOrDefault(item => item.Codec == codec);
            if (option is not null)
            {
                CodecComboBox.SelectedItem = option;
            }
        }

        if (payload.TryGetProperty("resolution", out JsonElement resolutionElement))
        {
            string? resolution = resolutionElement.GetString();
            ResolutionOption? option = ResolutionComboBox.Items.OfType<ResolutionOption>().FirstOrDefault(item => item.FilterExpression == resolution);
            if (option is not null)
            {
                ResolutionComboBox.SelectedItem = option;
            }
        }

        if (payload.TryGetProperty("preset", out JsonElement presetElement))
        {
            string? preset = presetElement.GetString();
            PresetOption? option = PresetComboBox.Items.OfType<PresetOption>().FirstOrDefault(item => item.Preset == preset);
            if (option is not null)
            {
                PresetComboBox.SelectedItem = option;
            }
        }

        if (payload.TryGetProperty("crf", out JsonElement crfElement) && crfElement.TryGetInt32(out int crf))
        {
            crf = Math.Clamp(crf, 0, 51);
            CrfSlider.Value = crf;
            CrfTextBox.Text = crf.ToString(CultureInfo.InvariantCulture);
        }

        if (payload.TryGetProperty("enableCrop", out JsonElement enableCropElement))
        {
            EnableCropCheckBox.IsChecked = enableCropElement.GetBoolean();
        }

        if (payload.TryGetProperty("cropX", out JsonElement cropXElement))
        {
            CropXTextBox.Text = cropXElement.GetString() ?? "0";
        }

        if (payload.TryGetProperty("cropY", out JsonElement cropYElement))
        {
            CropYTextBox.Text = cropYElement.GetString() ?? "0";
        }

        if (payload.TryGetProperty("cropW", out JsonElement cropWElement))
        {
            CropWTextBox.Text = cropWElement.GetString() ?? "0";
        }

        if (payload.TryGetProperty("cropH", out JsonElement cropHElement))
        {
            CropHTextBox.Text = cropHElement.GetString() ?? "0";
        }

        if (payload.TryGetProperty("disableAudio", out JsonElement disableAudioElement))
        {
            DisableAudioCheckBox.IsChecked = disableAudioElement.GetBoolean();
        }

        if (payload.TryGetProperty("overwrite", out JsonElement overwriteElement))
        {
            OverwriteCheckBox.IsChecked = overwriteElement.GetBoolean();
        }

        if (payload.TryGetProperty("hevcOnlyMode", out JsonElement hevcOnlyModeElement))
        {
            HevcOnlyModeCheckBox.IsChecked = hevcOnlyModeElement.GetBoolean();
        }

        if (payload.TryGetProperty("audioBitrate", out JsonElement audioBitrateElement))
        {
            string bitrate = audioBitrateElement.GetString() ?? "192k";
            if (AudioBitrateComboBox.Items.OfType<string>().Any(item => item == bitrate))
            {
                AudioBitrateComboBox.SelectedItem = bitrate;
            }
        }

        SaveSettings();
        NotifyWebModuleSnapshotChanged();
        return GetWebModuleSnapshot();
    }

    public object InvokeWebModuleAction(string action)
    {
        switch (action)
        {
            case "add-files":
                AddFilesButton_Click(this, new RoutedEventArgs());
                break;
            case "add-folder":
                AddFolderButton_Click(this, new RoutedEventArgs());
                break;
            case "remove-selected":
                RemoveSelectedButton_Click(this, new RoutedEventArgs());
                break;
            case "clear-queue":
                ClearButton_Click(this, new RoutedEventArgs());
                break;
            case "start-encoding":
                StartButton_Click(this, new RoutedEventArgs());
                break;
            case "cancel-encoding":
                CancelButton_Click(this, new RoutedEventArgs());
                break;
            case "browse-ffmpeg":
                BrowseFfmpegFolderButton_Click(this, new RoutedEventArgs());
                break;
            case "browse-output":
                BrowseOutputFolderButton_Click(this, new RoutedEventArgs());
                break;
        }

        return GetWebModuleSnapshot();
    }

    private static string ToDataUrl(BitmapSource bitmap)
    {
        using var stream = new MemoryStream();
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        encoder.Save(stream);
        return $"data:image/png;base64,{Convert.ToBase64String(stream.ToArray())}";
    }

    private void NotifyWebModuleSnapshotChanged()
    {
        WebModuleSnapshotChanged?.Invoke();
    }

    private static bool TryResolveFfmpegTools(string folderPath, out string ffmpegExe, out string ffprobeExe, out string errorMessage)
    {
        ffmpegExe = string.Empty;
        ffprobeExe = string.Empty;
        errorMessage = string.Empty;

        if (string.IsNullOrWhiteSpace(folderPath))
        {
            errorMessage = "Please select FFmpeg folder.";
            return false;
        }

        if (!Directory.Exists(folderPath))
        {
            errorMessage = "FFmpeg folder does not exist.";
            return false;
        }

        ffmpegExe = Path.Combine(folderPath, "ffmpeg.exe");
        ffprobeExe = Path.Combine(folderPath, "ffprobe.exe");

        if (!File.Exists(ffmpegExe) || !File.Exists(ffprobeExe))
        {
            errorMessage = "Selected folder must contain ffmpeg.exe and ffprobe.exe.";
            return false;
        }

        return true;
    }

    private static async Task<VideoProbeInfo?> ProbeVideoInfoAsync(string ffprobeExe, string inputPath, CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = ffprobeExe,
            Arguments = $"-v error -print_format json -show_format -show_streams \"{inputPath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        try
        {
            using var proc = new Process { StartInfo = psi };
            proc.Start();
            string output = await proc.StandardOutput.ReadToEndAsync();
            _ = await proc.StandardError.ReadToEndAsync();
            await proc.WaitForExitAsync(ct);

            if (proc.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
            {
                return null;
            }

            using JsonDocument doc = JsonDocument.Parse(output);

            JsonElement format = doc.RootElement.TryGetProperty("format", out JsonElement fmt) ? fmt : default;
            JsonElement streams = doc.RootElement.TryGetProperty("streams", out JsonElement strm) ? strm : default;

            JsonElement? videoStream = null;
            JsonElement? audioStream = null;
            if (streams.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement stream in streams.EnumerateArray())
                {
                    string? codecType = GetString(stream, "codec_type");
                    if (videoStream is null && string.Equals(codecType, "video", StringComparison.OrdinalIgnoreCase))
                    {
                        videoStream = stream;
                    }
                    else if (audioStream is null && string.Equals(codecType, "audio", StringComparison.OrdinalIgnoreCase))
                    {
                        audioStream = stream;
                    }
                }
            }

            double fps = ParseFraction(GetString(videoStream, "r_frame_rate"));
            return new VideoProbeInfo(
                ParseDouble(GetString(format, "duration")),
                ParseLong(GetString(format, "size")),
                ParseLong(GetString(format, "bit_rate")),
                GetString(videoStream, "codec_name"),
                ParseInt(GetString(videoStream, "width")),
                ParseInt(GetString(videoStream, "height")),
                fps,
                ParseLong(GetString(videoStream, "bit_rate")),
                GetString(audioStream, "codec_name"),
                ParseInt(GetString(audioStream, "sample_rate")),
                ParseInt(GetString(audioStream, "channels")),
                ParseLong(GetString(audioStream, "bit_rate")));
        }
        catch
        {
            return null;
        }
    }

    private static string? GetString(JsonElement? element, string propertyName)
    {
        if (element is null || element.Value.ValueKind == JsonValueKind.Undefined || element.Value.ValueKind == JsonValueKind.Null)
        {
            return null;
        }

        if (!element.Value.TryGetProperty(propertyName, out JsonElement value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.GetRawText(),
            _ => null
        };
    }

    private static int ParseInt(string? value) =>
        int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0;

    private static long ParseLong(string? value) =>
        long.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out long parsed) ? parsed : 0;

    private static double ParseDouble(string? value) =>
        double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed) ? parsed : 0;

    private static double ParseFraction(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }

        string[] parts = value.Split('/');
        if (parts.Length == 2
            && double.TryParse(parts[0], NumberStyles.Any, CultureInfo.InvariantCulture, out double numerator)
            && double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double denominator)
            && Math.Abs(denominator) > 0.000001)
        {
            return numerator / denominator;
        }

        return ParseDouble(value);
    }

    private static string FormatDuration(double seconds)
    {
        if (seconds <= 0)
        {
            return "-";
        }

        TimeSpan span = TimeSpan.FromSeconds(seconds);
        return span.TotalHours >= 1
            ? $"{(int)span.TotalHours:D2}:{span.Minutes:D2}:{span.Seconds:D2}"
            : $"{span.Minutes:D2}:{span.Seconds:D2}";
    }

    private static string FormatBytes(long bytes)
    {
        if (bytes <= 0)
        {
            return "-";
        }

        string[] units = ["B", "KB", "MB", "GB", "TB"];
        double size = bytes;
        int unit = 0;
        while (size >= 1024 && unit < units.Length - 1)
        {
            size /= 1024;
            unit++;
        }

        return $"{size:F1} {units[unit]}";
    }

    private static string FormatBitrate(long bitsPerSecond)
    {
        if (bitsPerSecond <= 0)
        {
            return "-";
        }

        return bitsPerSecond >= 1_000_000
            ? $"{bitsPerSecond / 1_000_000.0:F2} Mbps"
            : $"{bitsPerSecond / 1_000.0:F0} kbps";
    }

    private static string FormatFps(double fps) =>
        fps > 0 ? fps.ToString("F2", CultureInfo.InvariantCulture) : "-";

    private static bool IsHevcCodec(string? codecName) =>
        string.Equals(codecName, "hevc", StringComparison.OrdinalIgnoreCase)
        || string.Equals(codecName, "h265", StringComparison.OrdinalIgnoreCase)
        || string.Equals(codecName, "h.265", StringComparison.OrdinalIgnoreCase);

    private readonly record struct EncodingSettings(
        string FfmpegPath,
        string FfprobePath,
        string OutputFolder,
        string VideoCodec,
        string ResolutionFilterExpression,
        string Preset,
        int Crf,
        bool EnableCrop,
        int CropX,
        int CropY,
        int CropW,
        int CropH,
        bool DisableAudio,
        string AudioBitrate,
        bool Overwrite,
        bool HevcOnlyMode);

    private sealed class VideoConverterSettings
    {
        public string FfmpegFolder { get; set; } = string.Empty;
        public string OutputFolder { get; set; } = string.Empty;
        public bool HevcOnlyMode { get; set; } = true;
    }

    private readonly record struct VideoProbeInfo(
        double DurationSeconds,
        long FileSizeBytes,
        long FormatBitrate,
        string? VideoCodec,
        int Width,
        int Height,
        double Fps,
        long VideoBitrate,
        string? AudioCodec,
        int AudioSampleRate,
        int AudioChannels,
        long AudioBitrate);

    private sealed record CodecOption(string DisplayName, string Codec)
    {
        public override string ToString() => DisplayName;
    }

    private sealed record ResolutionOption(string DisplayName, string FilterExpression)
    {
        public override string ToString() => DisplayName;
    }

    private sealed record PresetOption(string DisplayName, string Preset)
    {
        public override string ToString() => DisplayName;
    }
}

public sealed class EncodingQueueItem : INotifyPropertyChanged
{
    private double _progress;
    private string _status = "Queued";
    private string _outputPath = string.Empty;
    private string _videoCodec = "-";
    private bool _isHevc;

    public EncodingQueueItem(string inputPath)
    {
        InputPath = inputPath;
    }

    public string InputPath { get; }

    public string FileName => Path.GetFileName(InputPath);

    public string OutputPath
    {
        get => _outputPath;
        set
        {
            if (value == _outputPath)
            {
                return;
            }

            _outputPath = value;
            OnPropertyChanged();
        }
    }

    public string Status
    {
        get => _status;
        set
        {
            if (value == _status)
            {
                return;
            }

            _status = value;
            OnPropertyChanged();
        }
    }

    public string? ErrorMessage { get; set; }

    public string VideoCodec
    {
        get => _videoCodec;
        set
        {
            if (value == _videoCodec)
            {
                return;
            }

            _videoCodec = value;
            OnPropertyChanged();
        }
    }

    public bool IsHevc
    {
        get => _isHevc;
        set
        {
            if (value == _isHevc)
            {
                return;
            }

            _isHevc = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsHevcDisplay));
        }
    }

    public string IsHevcDisplay => IsHevc ? "Yes" : "No";

    public double Progress
    {
        get => _progress;
        set
        {
            double clamped = Math.Clamp(value, 0, 100);
            if (Math.Abs(clamped - _progress) < 0.0001)
            {
                return;
            }

            _progress = clamped;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ProgressText));
        }
    }

    public string ProgressText => $"{Progress:F1}%";

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
