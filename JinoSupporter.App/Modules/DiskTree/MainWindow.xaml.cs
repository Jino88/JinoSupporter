using DiskTree.Models;
using DiskTree.Services;
using Microsoft.Win32;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Text.Json;
using WorkbenchHost.Infrastructure;

namespace DiskTree;

public partial class MainWindow : Window
{
    public static string DefaultDatabasePath =>
        AppSettingsPathManager.GetModuleFilePath("DiskTree", "disktree-index.db");
    private const double ScanPhaseStartPercent = 8.0;
    private const double ScanPhaseEndPercent = 82.0;
    private const double DbPhaseStartPercent = 82.0;
    private const double DbPhaseEndPercent = 100.0;

    private readonly ObservableCollection<DiskNode> _rootNodes = new();
    private readonly ObservableCollection<DuplicateMatchRecord> _duplicateMatches = new();
    private readonly DiskScanner _scanner = new();
    private bool _isBusy;
    private string? _selectedHashSourceFilePath;
    public event Action? WebModuleSnapshotChanged;

    public MainWindow()
    {
        InitializeComponent();

        DiskTreeView.ItemsSource = _rootNodes;
        DuplicateDataGrid.ItemsSource = _duplicateMatches;
        UpdateDatabaseInfo(LoadCurrentIndexedRowCount());
        SelectedFileInfoTextBlock.Text = "Select a file in the left tree to check identical hashes.";
        UpdateCollectButtonState();
    }

    private void BrowseRootFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        var dialog = new OpenFolderDialog
        {
            Title = "Select root folder",
            Multiselect = false
        };

        if (dialog.ShowDialog() == true)
        {
            RootFolderPathTextBox.Text = dialog.FolderName;
            NotifyWebModuleSnapshotChanged();
        }
    }

    private async void ScanAndIndexButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        string rootPath = RootFolderPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
        {
            MessageBox.Show("Select a valid root folder first.", "Invalid Folder", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetBusyState(true);
        ClearDuplicateSelectionState("No file comparison yet.");

        try
        {
            OperationProgressBar.IsIndeterminate = true;
            OperationProgressBar.Value = 0;
            StatusTextBlock.Text = "Estimating file count...";

            var estimateProgress = new Progress<string>(status => StatusTextBlock.Text = status);
            int estimatedFileCount = await Task.Run(() => _scanner.EstimateFileCount(rootPath, CancellationToken.None, estimateProgress));

            if (estimatedFileCount <= 0)
            {
                _rootNodes.Clear();
                TreeSummaryTextBlock.Text = "No readable files found under the selected root.";

                int deletedCount = await Task.Run(() =>
                {
                    using var store = new DiskIndexStore(DefaultDatabasePath);
                    return store.DeleteMissingPathsUnderRoot(
                        rootPath,
                        new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                        CancellationToken.None);
                });

                int remainingRowCount = await Task.Run(() =>
                {
                    using var store = new DiskIndexStore(DefaultDatabasePath);
                    return store.GetIndexedFileCount();
                });

                UpdateDatabaseInfo(remainingRowCount);
                StatusTextBlock.Text = $"No files found. Removed {deletedCount:N0} stale DB rows.";
                OperationProgressBar.IsIndeterminate = false;
                OperationProgressBar.Value = 100;
                NotifyWebModuleSnapshotChanged();
                return;
            }

            StatusTextBlock.Text = "Loading existing index metadata...";
            IReadOnlyDictionary<string, IndexedFileSnapshot> existingIndex = await Task.Run(() =>
            {
                using var store = new DiskIndexStore(DefaultDatabasePath);
                return store.GetIndexedMetadataByRoot(rootPath);
            });
            bool isIncrementalMode = existingIndex.Count > 0;

            OperationProgressBar.IsIndeterminate = false;
            OperationProgressBar.Value = ScanPhaseStartPercent;
            StatusTextBlock.Text = isIncrementalMode
                ? $"Estimated files: {estimatedFileCount:N0}. Incremental scan + selective hashing..."
                : $"Estimated files: {estimatedFileCount:N0}. Initial scan + hashing...";

            var scanProgress = new Progress<ScanProgress>(update =>
            {
                double ratio = update.TotalFiles <= 0
                    ? 1.0
                    : Math.Min(1.0, (double)update.ProcessedFiles / update.TotalFiles);
                OperationProgressBar.Value = ScanPhaseStartPercent + ratio * (ScanPhaseEndPercent - ScanPhaseStartPercent);
                StatusTextBlock.Text = update.StatusText;
            });

            ScanResult scanResult = await Task.Run(() =>
                _scanner.Scan(rootPath, estimatedFileCount, CancellationToken.None, existingIndex, scanProgress));

            _rootNodes.Clear();
            _rootNodes.Add(scanResult.RootNode);
            ExpandRootNode();

            List<IndexedFileRecord> changedFiles = BuildChangedFiles(scanResult.IndexedFiles, existingIndex);
            var currentPathSet = scanResult.IndexedFiles
                .Select(file => file.FilePath)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            TreeSummaryTextBlock.Text =
                $"Root size: {scanResult.RootNode.SizeText} | " +
                $"Estimated files: {estimatedFileCount:N0} | " +
                $"Indexed files: {scanResult.ScannedFileCount:N0} | " +
                $"Rehashed/Reused: {scanResult.RehashedCount:N0}/{scanResult.ReusedHashCount:N0} | " +
                $"DB changes: {changedFiles.Count:N0} | " +
                $"Visited folders: {scanResult.ScannedDirectoryCount:N0} | " +
                $"Skipped folders/files: {scanResult.SkippedDirectoryCount:N0}/{scanResult.SkippedFileCount:N0}";

            StatusTextBlock.Text = "Applying changes to SQLite index...";
            OperationProgressBar.Value = DbPhaseStartPercent;

            var dbProgress = new Progress<IndexWriteProgress>(update =>
            {
                double ratio = update.TotalCount <= 0
                    ? 1.0
                    : Math.Min(1.0, (double)update.WrittenCount / update.TotalCount);
                OperationProgressBar.Value = DbPhaseStartPercent + ratio * (DbPhaseEndPercent - DbPhaseStartPercent);
                StatusTextBlock.Text = $"Updating DB rows: {update.WrittenCount:N0}/{Math.Max(update.TotalCount, update.WrittenCount):N0}";
            });

            int deletedRows = await Task.Run(() =>
            {
                using var store = new DiskIndexStore(DefaultDatabasePath);
                store.UpsertFiles(
                    changedFiles,
                    CancellationToken.None,
                    changedFiles.Count,
                    dbProgress);

                return store.DeleteMissingPathsUnderRoot(rootPath, currentPathSet, CancellationToken.None);
            });

            int indexedRowCount = await Task.Run(() =>
            {
                using var store = new DiskIndexStore(DefaultDatabasePath);
                return store.GetIndexedFileCount();
            });

            TreeSummaryTextBlock.Text += $" | Removed from DB: {deletedRows:N0}";
            UpdateDatabaseInfo(indexedRowCount);
            OperationProgressBar.IsIndeterminate = false;
            OperationProgressBar.Value = 100;
            StatusTextBlock.Text =
                $"Update complete. Scanned {scanResult.ScannedFileCount:N0} files | Changed {changedFiles.Count:N0} | Deleted {deletedRows:N0} | DB rows {indexedRowCount:N0}.";
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Scan failed.";
            MessageBox.Show($"Scan/index failed:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            NotifyWebModuleSnapshotChanged();
        }
        finally
        {
            SetBusyState(false);
        }
    }

    private void BrowseCompareFileButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "Select file to compare",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            CompareFilePathTextBox.Text = dialog.FileName;
            NotifyWebModuleSnapshotChanged();
        }
    }

    private async void DiskTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (_isBusy)
        {
            return;
        }

        if (e.NewValue is not DiskNode node || node.IsDirectory)
        {
            return;
        }

        if (!File.Exists(node.FullPath))
        {
            return;
        }

        CompareFilePathTextBox.Text = node.FullPath;
        await FindIdenticalFilesForPathAsync(node.FullPath, triggeredByTreeSelection: true);
    }

    private async void FindIdenticalFilesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        string filePath = CompareFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            MessageBox.Show("Select a valid file first.", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        await FindIdenticalFilesForPathAsync(filePath, triggeredByTreeSelection: false);
    }

    private async Task FindIdenticalFilesForPathAsync(string filePath, bool triggeredByTreeSelection)
    {
        if (!File.Exists(DefaultDatabasePath))
        {
            MessageBox.Show("Index DB does not exist yet. Run Scan + Update first.", "Index Missing", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        SetBusyState(true);

        try
        {
            _duplicateMatches.Clear();
            DuplicateSummaryTextBlock.Text = triggeredByTreeSelection
                ? "Searching identical hashes from selected tree file..."
                : "Searching candidates...";
            StatusTextBlock.Text = "Computing quick hash for selected file...";

            var selectedFileInfo = new FileInfo(filePath);
            string selectedFullPath = selectedFileInfo.FullName;
            long selectedFileSize = selectedFileInfo.Length;

            string selectedPartialHash = await Task.Run(() =>
                FileHasher.ComputeHeadTailHash(selectedFullPath, selectedFileSize));
            _selectedHashSourceFilePath = selectedFullPath;
            SelectedFileInfoTextBlock.Text =
                $"Selected: {selectedFullPath}\nHead/Tail SHA-256: {selectedPartialHash}\nFull SHA-256: (verifying...)";

            IReadOnlyList<IndexedFileRecord> candidates = await Task.Run(() =>
            {
                using var store = new DiskIndexStore(DefaultDatabasePath);
                return store.FindCandidates(selectedFileSize, selectedPartialHash, selectedFullPath);
            });

            if (candidates.Count == 0)
            {
                int indexedRowCount = await Task.Run(() =>
                {
                    using var store = new DiskIndexStore(DefaultDatabasePath);
                    return store.GetIndexedFileCount();
                });

                UpdateDatabaseInfo(indexedRowCount);
                DuplicateSummaryTextBlock.Text = $"No candidates found in current index (DB rows: {indexedRowCount:N0}).";
                StatusTextBlock.Text = "No identical files found.";
                SelectedFileInfoTextBlock.Text =
                    $"Selected: {selectedFullPath}\nHead/Tail SHA-256: {selectedPartialHash}\nFull SHA-256: (no duplicates)";
                UpdateCollectButtonState();
                NotifyWebModuleSnapshotChanged();
                return;
            }

            DuplicateSummaryTextBlock.Text = $"Candidates from index: {candidates.Count:N0}. Verifying full hash...";
            StatusTextBlock.Text = "Verifying exact matches with full SHA-256 hash...";

            string selectedFullHash = await Task.Run(() => FileHasher.ComputeFullHash(selectedFullPath));

            IReadOnlyList<DuplicateMatchRecord> verifiedMatches = await Task.Run(() =>
            {
                var matches = new List<DuplicateMatchRecord>();
                foreach (IndexedFileRecord candidate in candidates)
                {
                    if (!File.Exists(candidate.FilePath))
                    {
                        continue;
                    }

                    try
                    {
                        string candidateFullHash = FileHasher.ComputeFullHash(candidate.FilePath);
                        if (string.Equals(candidateFullHash, selectedFullHash, StringComparison.OrdinalIgnoreCase))
                        {
                            matches.Add(new DuplicateMatchRecord(candidate.FilePath, candidate.FileSize, candidate.LastWriteUtc));
                        }
                    }
                    catch
                    {
                        // Skip unreadable files while validating duplicates.
                    }
                }

                return (IReadOnlyList<DuplicateMatchRecord>)matches;
            });

            foreach (DuplicateMatchRecord match in verifiedMatches)
            {
                _duplicateMatches.Add(match);
            }

            DuplicateSummaryTextBlock.Text = $"Exact identical files: {_duplicateMatches.Count:N0}";
            StatusTextBlock.Text = _duplicateMatches.Count == 0
                ? "No exact identical files found."
                : $"Found {_duplicateMatches.Count:N0} exact identical files.";

            SelectedFileInfoTextBlock.Text =
                $"Selected: {selectedFullPath}\nHead/Tail SHA-256: {selectedPartialHash}\nFull SHA-256: {selectedFullHash}";
            UpdateCollectButtonState();
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "File comparison failed.";
            MessageBox.Show($"Duplicate search failed:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            NotifyWebModuleSnapshotChanged();
        }
        finally
        {
            SetBusyState(false);
        }
    }

    private async void CollectMatchesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(_selectedHashSourceFilePath) || !File.Exists(_selectedHashSourceFilePath))
        {
            MessageBox.Show("Select a valid file first.", "No Source File", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (_duplicateMatches.Count == 0)
        {
            MessageBox.Show("No identical files to collect yet.", "No Matches", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        SetBusyState(true);

        try
        {
            StatusTextBlock.Text = "Collecting identical files into temporary folder...";

            (string targetFolder, int copiedCount, int skippedCount) = await Task.Run(() =>
                CollectCurrentMatchesToTemp(_selectedHashSourceFilePath, _duplicateMatches));

            StatusTextBlock.Text = $"Collected {copiedCount:N0} files to temp folder (skipped {skippedCount:N0}).";
            DuplicateSummaryTextBlock.Text = $"Collected to: {targetFolder}";

            Process.Start(new ProcessStartInfo
            {
                FileName = targetFolder,
                UseShellExecute = true
            });
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Collect failed.";
            MessageBox.Show($"Collect failed:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            NotifyWebModuleSnapshotChanged();
        }
        finally
        {
            SetBusyState(false);
            UpdateCollectButtonState();
        }
    }

    private void SetBusyState(bool isBusy)
    {
        _isBusy = isBusy;

        BrowseRootFolderButton.IsEnabled = !isBusy;
        ScanAndIndexButton.IsEnabled = !isBusy;
        BrowseCompareFileButton.IsEnabled = !isBusy;
        FindIdenticalFilesButton.IsEnabled = !isBusy;
        DiskTreeView.IsEnabled = !isBusy;

        OperationProgressBar.IsIndeterminate = false;
        OperationProgressBar.Value = isBusy ? 0 : 100;
        UpdateCollectButtonState();
        NotifyWebModuleSnapshotChanged();
    }

    private void ExpandRootNode()
    {
        ExpandRootNode(attempt: 0);
    }

    private void ExpandRootNode(int attempt)
    {
        Dispatcher.BeginInvoke(() =>
        {
            if (_rootNodes.Count == 0)
            {
                return;
            }

            if (DiskTreeView.ItemContainerGenerator.ContainerFromItem(_rootNodes[0]) is TreeViewItem rootTreeItem)
            {
                rootTreeItem.IsExpanded = true;
                return;
            }

            if (attempt < 6)
            {
                ExpandRootNode(attempt + 1);
            }
        }, DispatcherPriority.Loaded);
    }

    private void UpdateDatabaseInfo(int? indexedRowCount)
    {
        if (indexedRowCount is int rowCount)
        {
            DatabasePathTextBlock.Text = $"SQLite DB: {DefaultDatabasePath} | Rows: {rowCount:N0}";
            return;
        }

        DatabasePathTextBlock.Text = $"SQLite DB: {DefaultDatabasePath}";
    }

    private int? LoadCurrentIndexedRowCount()
    {
        if (!File.Exists(DefaultDatabasePath))
        {
            return null;
        }

        try
        {
            using var store = new DiskIndexStore(DefaultDatabasePath);
            return store.GetIndexedFileCount();
        }
        catch
        {
            return null;
        }
    }

    private void ClearDuplicateSelectionState(string summaryText)
    {
        _duplicateMatches.Clear();
        _selectedHashSourceFilePath = null;
        DuplicateSummaryTextBlock.Text = summaryText;
        SelectedFileInfoTextBlock.Text = "Select a file in the left tree to check identical hashes.";
        UpdateCollectButtonState();
        NotifyWebModuleSnapshotChanged();
    }

    private void UpdateCollectButtonState()
    {
        bool hasSourceFile = !string.IsNullOrWhiteSpace(_selectedHashSourceFilePath) && File.Exists(_selectedHashSourceFilePath);
        CollectMatchesButton.IsEnabled = !_isBusy && hasSourceFile && _duplicateMatches.Count > 0;
    }

    private static (string TargetFolder, int CopiedCount, int SkippedCount) CollectCurrentMatchesToTemp(
        string sourceFilePath,
        IEnumerable<DuplicateMatchRecord> matches)
    {
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string sourceName = Path.GetFileNameWithoutExtension(sourceFilePath);
        string safeSourceName = MakeSafeFileName(string.IsNullOrWhiteSpace(sourceName) ? "selected" : sourceName);

        string targetFolder = Path.Combine(
            Path.GetTempPath(),
            "DataGraphHost",
            "DiskTreeCollect",
            $"{timestamp}_{safeSourceName}");

        Directory.CreateDirectory(targetFolder);

        var orderedPaths = new List<string> { sourceFilePath };
        orderedPaths.AddRange(matches.Select(item => item.FilePath));
        List<string> uniquePaths = orderedPaths
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var mapLines = new List<string>
        {
            $"CollectedAt\t{DateTime.Now:O}",
            "CollectedName\tSourcePath"
        };

        int copiedCount = 0;
        int skippedCount = 0;

        for (int index = 0; index < uniquePaths.Count; index++)
        {
            string sourcePath = uniquePaths[index];
            if (!File.Exists(sourcePath))
            {
                skippedCount++;
                continue;
            }

            string sourceFileName = Path.GetFileName(sourcePath);
            string safeFileName = MakeSafeFileName(string.IsNullOrWhiteSpace(sourceFileName)
                ? $"file_{index + 1:D4}"
                : sourceFileName);

            string targetFileName = $"{index + 1:D4}_{safeFileName}";
            string targetPath = Path.Combine(targetFolder, targetFileName);

            while (File.Exists(targetPath))
            {
                targetFileName = $"{index + 1:D4}_{Guid.NewGuid():N}_{safeFileName}";
                targetPath = Path.Combine(targetFolder, targetFileName);
            }

            try
            {
                File.Copy(sourcePath, targetPath, overwrite: false);
                mapLines.Add($"{targetFileName}\t{sourcePath}");
                copiedCount++;
            }
            catch
            {
                skippedCount++;
            }
        }

        File.WriteAllLines(Path.Combine(targetFolder, "_source_map.tsv"), mapLines);
        return (targetFolder, copiedCount, skippedCount);
    }

    private static string MakeSafeFileName(string fileName)
    {
        char[] invalidChars = Path.GetInvalidFileNameChars();
        var safeChars = fileName
            .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
            .ToArray();
        return new string(safeChars);
    }

    private static List<IndexedFileRecord> BuildChangedFiles(
        IReadOnlyList<IndexedFileRecord> scannedFiles,
        IReadOnlyDictionary<string, IndexedFileSnapshot> existingIndex)
    {
        var changedFiles = new List<IndexedFileRecord>(capacity: scannedFiles.Count);

        foreach (IndexedFileRecord scannedFile in scannedFiles)
        {
            if (!existingIndex.TryGetValue(scannedFile.FilePath, out IndexedFileSnapshot existing))
            {
                changedFiles.Add(scannedFile);
                continue;
            }

            bool isUnchanged =
                existing.FileSize == scannedFile.FileSize &&
                existing.LastWriteUtc == scannedFile.LastWriteUtc &&
                string.Equals(existing.HeadTailHash, scannedFile.HeadTailHash, StringComparison.OrdinalIgnoreCase);

            if (!isUnchanged)
            {
                changedFiles.Add(scannedFile);
            }
        }

        return changedFiles;
    }

    public object GetWebModuleSnapshot()
    {
        return new
        {
            moduleType = "DiskTree",
            rootFolderPath = RootFolderPathTextBox.Text ?? string.Empty,
            compareFilePath = CompareFilePathTextBox.Text ?? string.Empty,
            statusMessage = StatusTextBlock.Text ?? string.Empty,
            treeSummary = TreeSummaryTextBlock.Text ?? string.Empty,
            duplicateSummary = DuplicateSummaryTextBlock.Text ?? string.Empty,
            selectedFileInfo = SelectedFileInfoTextBlock.Text ?? string.Empty,
            databaseInfo = DatabasePathTextBlock.Text ?? string.Empty,
            progressValue = OperationProgressBar.Value,
            isBusy = _isBusy,
            canCollect = CollectMatchesButton.IsEnabled,
            rootNodes = _rootNodes.Take(40).Select(node => BuildNodeSnapshot(node, 0)).ToArray(),
            duplicateMatches = _duplicateMatches.Take(100).Select(match => new
            {
                filePath = match.FilePath,
                sizeText = match.SizeText,
                lastWriteText = match.LastWriteText
            }).ToArray()
        };
    }

    public object UpdateWebModuleState(JsonElement payload)
    {
        return GetWebModuleSnapshot();
    }

    public object InvokeWebModuleAction(string action)
    {
        switch (action)
        {
            case "browse-root-folder":
                BrowseRootFolderButton_Click(this, new RoutedEventArgs());
                break;
            case "scan-and-update":
                ScanAndIndexButton_Click(this, new RoutedEventArgs());
                break;
            case "browse-compare-file":
                BrowseCompareFileButton_Click(this, new RoutedEventArgs());
                break;
            case "find-identical":
                FindIdenticalFilesButton_Click(this, new RoutedEventArgs());
                break;
            case "collect-matches":
                CollectMatchesButton_Click(this, new RoutedEventArgs());
                break;
        }

        return GetWebModuleSnapshot();
    }

    private static object BuildNodeSnapshot(DiskNode node, int depth)
    {
        return new
        {
            name = node.Name,
            kindText = node.KindText,
            sizeText = node.SizeText,
            percentText = node.PercentText,
            percentOfRoot = node.PercentOfRoot,
            depth,
            children = node.Children.Take(24).Select(child => BuildNodeSnapshot(child, depth + 1)).ToArray()
        };
    }

    private void NotifyWebModuleSnapshotChanged()
    {
        WebModuleSnapshotChanged?.Invoke();
    }
}
