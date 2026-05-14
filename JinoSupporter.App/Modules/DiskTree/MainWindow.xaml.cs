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
    private static string LastRootSettingsPath =>
        AppSettingsPathManager.GetModuleFilePath("DiskTree", "last-root.txt");
    private static string LastTempFolderSettingsPath =>
        AppSettingsPathManager.GetModuleFilePath("DiskTree", "last-temp.txt");
    private static string DefaultTempFolderPath =>
        Path.Combine(Path.GetTempPath(), "DataGraphHost", "DiskTreeCollect");
    private const double ScanPhaseStartPercent = 8.0;
    private const double ScanPhaseEndPercent = 82.0;
    private const double DbPhaseStartPercent = 82.0;
    private const double DbPhaseEndPercent = 100.0;

    private readonly ObservableCollection<DiskNode> _rootNodes = new();
    private readonly ObservableCollection<DuplicateMatchRecord> _duplicateMatches = new();
    private Dictionary<int, IReadOnlyList<DuplicateMatchRecord>> _groupMembersById = new();
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

        // Restore the saved Temp folder, falling back to the system temp default on first run
        // so Collect-To-Temp always has a valid destination ready to go.
        TempFolderPathTextBox.Text = LoadLastTempPath() ?? DefaultTempFolderPath;

        // Restore the last-used root so the panes can auto-populate from the existing index
        // without forcing the user to re-pick + re-scan every time the window opens.
        string? lastRoot = LoadLastRootPath();
        if (!string.IsNullOrWhiteSpace(lastRoot))
        {
            RootFolderPathTextBox.Text = lastRoot;
            Dispatcher.BeginInvoke(async () =>
            {
                if (File.Exists(DefaultDatabasePath) && Directory.Exists(lastRoot))
                {
                    await LoadTreeFromIndexAsync(lastRoot);
                    await LoadDuplicateGroupsInScopeAsync(lastRoot);
                }
            }, DispatcherPriority.Loaded);
        }
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
            SaveLastRootPath(dialog.FolderName);
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void BrowseTempFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy) return;

        var dialog = new OpenFolderDialog
        {
            Title = "Select temp folder for collected duplicates",
            Multiselect = false
        };

        if (dialog.ShowDialog() == true)
        {
            TempFolderPathTextBox.Text = dialog.FolderName;
            SaveLastTempPath(dialog.FolderName);
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

            // Auto-populate the duplicate group list now that the index is up to date.
            await LoadDuplicateGroupsInScopeAsync(rootPath);
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
        if (_isBusy) return;
        if (e.NewValue is not DiskNode node) return;

        // Tree selection drives a folder-scoped duplicate scan.
        // - Folder selected → scope = that folder (its subtree).
        // - File   selected → scope = its parent folder.
        string? scopePath = node.IsDirectory
            ? node.FullPath
            : Path.GetDirectoryName(node.FullPath);
        if (string.IsNullOrEmpty(scopePath) || !Directory.Exists(scopePath)) return;

        CompareFilePathTextBox.Text = scopePath;   // hidden control kept for web-action compat
        await LoadDuplicateGroupsInScopeAsync(scopePath);
    }

    private async Task LoadDuplicateGroupsInScopeAsync(string scopePath)
    {
        if (!File.Exists(DefaultDatabasePath))
        {
            MessageBox.Show("Index DB does not exist yet. Run Scan + Update first.", "Index Missing",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        SetBusyState(true);
        try
        {
            ResetPaneSources();
            _selectedHashSourceFilePath = null;
            AllFilesSummaryTextBlock.Text  = $"Scanning files under {scopePath} …";
            DuplicateSummaryTextBlock.Text = "Computing duplicates…";
            StatusTextBlock.Text           = "Loading files and detecting duplicates…";

            // One DB read → derive all three panes:
            //   1. allFilesView    = every indexed file (left pane)
            //   2. duplicatesView  = subset that belongs to a group of size > 1 (middle pane)
            //   3. groupMembersById = group-id → all members (used to populate right pane on selection)
            var (allFilesView, duplicatesView, groupMembersById) = await Task.Run(() =>
            {
                using var store = new DiskIndexStore(DefaultDatabasePath);
                IReadOnlyList<IndexedFileRecord> rawFiles = store.ListFilesUnder(scopePath);

                // Drop Recycle Bin entries ($Recycle.Bin, RECYCLER, etc.) — they're per-user trash and
                // pollute the duplicate list. Match anywhere in the path so nested mounts are covered too.
                var files = rawFiles
                    .Where(f => f.FilePath.IndexOf("Recycle", StringComparison.OrdinalIgnoreCase) < 0)
                    .ToList();

                // Group by (size + head/tail hash). size+head/tail SHA match is an effectively
                // collision-free duplicate signal in practice. Files with size 0 are skipped.
                var pathToGroup = new Dictionary<string, (int GroupId, int GroupSize)>(StringComparer.OrdinalIgnoreCase);
                int nextGid = 0;
                foreach (var grouping in files
                    .Where(f => f.FileSize > 0)
                    .GroupBy(f => (f.FileSize, f.HeadTailHash)))
                {
                    var members = grouping.ToList();
                    if (members.Count <= 1) continue;
                    nextGid++;
                    foreach (IndexedFileRecord f in members)
                    {
                        pathToGroup[f.FilePath] = (nextGid, members.Count);
                    }
                }

                var allView = new List<DuplicateMatchRecord>(files.Count);
                var dupView = new List<DuplicateMatchRecord>();
                var byId    = new Dictionary<int, List<DuplicateMatchRecord>>();

                foreach (IndexedFileRecord f in files.OrderBy(x => x.FilePath, StringComparer.OrdinalIgnoreCase))
                {
                    pathToGroup.TryGetValue(f.FilePath, out var meta);
                    var record = new DuplicateMatchRecord(
                        f.FilePath, f.FileSize, f.LastWriteUtc, meta.GroupId, meta.GroupSize);
                    allView.Add(record);

                    if (meta.GroupId > 0)
                    {
                        if (!byId.TryGetValue(meta.GroupId, out var bucket))
                        {
                            bucket = new List<DuplicateMatchRecord>();
                            byId[meta.GroupId] = bucket;
                        }
                        bucket.Add(record);
                    }
                }

                // ── Folder-level duplicate detection ───────────────────────────────────────────
                // Detect "this whole folder is duplicated as that whole folder" by fingerprinting
                // each directory bottom-up (sorted child-files + sorted child-dir fingerprints).
                // Folders match when (last-segment name) AND (recursive content fingerprint) are equal.
                // Folder rows get group-ids in a high-numbered range so they don't collide with file ids.
                const int FolderIdBase = 1_000_000;
                var folderGroups = ComputeDuplicateFolderGroups(files);

                int folderGid = 0;
                var folderRecords = new List<DuplicateMatchRecord>();
                foreach (var (memberPaths, memberSize, memberLastWrite) in folderGroups)
                {
                    folderGid++;
                    int gid = FolderIdBase + folderGid;
                    int gsize = memberPaths.Count;

                    var bucket = new List<DuplicateMatchRecord>(gsize);
                    for (int i = 0; i < memberPaths.Count; i++)
                    {
                        var rec = new DuplicateMatchRecord(
                            memberPaths[i], memberSize[i], memberLastWrite[i], gid, gsize, IsDirectory: true);
                        bucket.Add(rec);
                        folderRecords.Add(rec);
                    }
                    byId[gid] = bucket;
                }

                // Middle pane: folder rows on top (largest folder dup first), then file rows by size desc.
                var fileRows = byId
                    .Where(kv => kv.Key < FolderIdBase)
                    .SelectMany(kv => kv.Value)
                    .OrderByDescending(record => record.FileSize)
                    .ThenBy(record => record.FilePath, StringComparer.OrdinalIgnoreCase);

                var folderRows = folderRecords
                    .OrderByDescending(record => record.FileSize)
                    .ThenBy(record => record.FilePath, StringComparer.OrdinalIgnoreCase);

                dupView = folderRows.Concat(fileRows).ToList();

                var byIdReadOnly = byId.ToDictionary(
                    kv => kv.Key,
                    kv => (IReadOnlyList<DuplicateMatchRecord>)kv.Value);

                return (allView, dupView, byIdReadOnly);
            });

            // Bulk-assign for performance — a 100k-row ObservableCollection.Add loop is unusable.
            _groupMembersById = groupMembersById;
            DuplicateFilesDataGrid.ItemsSource = duplicatesView;

            int folderGroupCount = groupMembersById.Count(kv => kv.Key >= 1_000_000);
            int fileGroupCount   = groupMembersById.Count - folderGroupCount;
            int totalDups        = duplicatesView.Count;
            AllFilesSummaryTextBlock.Text = $"{allFilesView.Count:N0} files indexed under {scopePath}";
            DuplicateSummaryTextBlock.Text = (fileGroupCount + folderGroupCount) == 0
                ? "No duplicates found in this scope."
                : folderGroupCount > 0
                    ? $"{folderGroupCount:N0} duplicate FOLDER group(s) on top, then {fileGroupCount:N0} file group(s). Click a row to see its duplicates →"
                    : $"{totalDups:N0} files in {fileGroupCount:N0} duplicate group(s). Click a row to see its duplicates →";
            StatusTextBlock.Text = (fileGroupCount + folderGroupCount) == 0
                ? "No duplicates."
                : $"Found {fileGroupCount:N0} file group(s) + {folderGroupCount:N0} folder group(s).";

            int indexedRowCount = await Task.Run(() =>
            {
                using var store = new DiskIndexStore(DefaultDatabasePath);
                return store.GetIndexedFileCount();
            });
            UpdateDatabaseInfo(indexedRowCount);
            UpdateCollectButtonState();
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            DuplicateSummaryTextBlock.Text = $"Failed: {ex.Message}";
            StatusTextBlock.Text            = "Duplicate scan failed.";
        }
        finally
        {
            SetBusyState(false);
        }
    }

    private void DuplicateFilesDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _duplicateMatches.Clear();

        if (DuplicateFilesDataGrid.SelectedItem is not DuplicateMatchRecord selected || selected.GroupId <= 0)
        {
            _selectedHashSourceFilePath = null;
            UpdateCollectButtonState();
            return;
        }

        if (_groupMembersById.TryGetValue(selected.GroupId, out IReadOnlyList<DuplicateMatchRecord>? members))
        {
            foreach (DuplicateMatchRecord member in members)
            {
                _duplicateMatches.Add(member);
            }
        }

        _selectedHashSourceFilePath = selected.FilePath;
        UpdateCollectButtonState();
        NotifyWebModuleSnapshotChanged();
    }

    /// <summary>Right-click on a DataGridRow normally doesn't change selection — promote it manually
    /// so the context-menu handlers can rely on DuplicateDataGrid.SelectedItem.</summary>
    private void DuplicateDataGridRow_PreviewMouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is DataGridRow row)
        {
            row.IsSelected = true;
        }
    }

    private void OpenInExplorerMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (DuplicateDataGrid.SelectedItem is not DuplicateMatchRecord rec) return;
        OpenRecordInExplorer(rec);
    }

    private void CopyPathMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (DuplicateDataGrid.SelectedItem is not DuplicateMatchRecord rec) return;
        try
        {
            Clipboard.SetText(rec.FilePath);
        }
        catch
        {
            // Clipboard access can fail transiently (another process holding it). Best-effort.
        }
    }

    private async void DeleteSelectedMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy) return;
        if (DuplicateDataGrid.SelectedItem is not DuplicateMatchRecord rec) return;

        SetBusyState(true);
        try
        {
            int dbRowsRemoved = 0;
            string targetPath = rec.FilePath;
            bool isDirectory = rec.IsDirectory;

            StatusTextBlock.Text = isDirectory
                ? $"Sending folder to Recycle Bin: {targetPath}"
                : $"Sending file to Recycle Bin: {targetPath}";

            await Task.Run(() =>
            {
                if (isDirectory)
                {
                    if (!Directory.Exists(targetPath)) return;

                    // Snapshot every indexed file under this folder BEFORE the OS delete so we can
                    // drop them from the index DB afterwards. Using ListFilesUnder = same scoping
                    // logic the duplicate panes already use, so the cleanup matches what the user sees.
                    using var store = new DiskIndexStore(DefaultDatabasePath);
                    var pathsUnder = store.ListFilesUnder(targetPath).Select(f => f.FilePath).ToList();

                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(
                        targetPath,
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);

                    if (pathsUnder.Count > 0)
                    {
                        dbRowsRemoved = store.DeleteFiles(pathsUnder);
                    }
                }
                else
                {
                    if (!File.Exists(targetPath)) return;

                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(
                        targetPath,
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);

                    using var store = new DiskIndexStore(DefaultDatabasePath);
                    dbRowsRemoved = store.DeleteFiles(new[] { targetPath });
                }
            });

            StatusTextBlock.Text = isDirectory
                ? $"Folder sent to Recycle Bin. Removed {dbRowsRemoved:N0} index rows."
                : $"File sent to Recycle Bin. Removed {dbRowsRemoved:N0} index row(s).";

            // Refresh tree + middle pane so the deleted item disappears from both views.
            string rootPath = RootFolderPathTextBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(rootPath) && Directory.Exists(rootPath))
            {
                await LoadTreeFromIndexAsync(rootPath);
                await LoadDuplicateGroupsInScopeAsync(rootPath);
            }
            else
            {
                _duplicateMatches.Clear();
                _selectedHashSourceFilePath = null;
            }

            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Delete failed.";
            MessageBox.Show($"Delete failed:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetBusyState(false);
            UpdateCollectButtonState();
        }
    }

    private static void OpenRecordInExplorer(DuplicateMatchRecord rec)
    {
        try
        {
            if (rec.IsDirectory)
            {
                if (Directory.Exists(rec.FilePath))
                {
                    Process.Start(new ProcessStartInfo { FileName = rec.FilePath, UseShellExecute = true });
                }
                return;
            }

            // For files: open Explorer with the file pre-selected. /select needs the path quoted
            // and uses comma as the argument separator — explorer is picky about formatting.
            if (File.Exists(rec.FilePath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName  = "explorer.exe",
                    Arguments = $"/select,\"{rec.FilePath}\"",
                    UseShellExecute = true
                });
                return;
            }

            // File no longer there (e.g., already moved) — fall back to its parent folder.
            string? parent = Path.GetDirectoryName(rec.FilePath);
            if (!string.IsNullOrEmpty(parent) && Directory.Exists(parent))
            {
                Process.Start(new ProcessStartInfo { FileName = parent, UseShellExecute = true });
            }
        }
        catch
        {
            // Best-effort: don't surface explorer launch failures.
        }
    }

    private void ResetPaneSources()
    {
        // Tree (left pane) is intentionally not cleared — it survives middle-pane re-scoping.
        DuplicateFilesDataGrid.ItemsSource = null;
        _duplicateMatches.Clear();
        _groupMembersById = new Dictionary<int, IReadOnlyList<DuplicateMatchRecord>>();
    }

    /// <summary>Rebuilds the folder-explorer tree (left pane) from the existing SQLite index,
    /// so reopening the window after a prior scan restores the navigable hierarchy without rescanning.
    /// Recycle Bin entries are skipped to keep the tree clean.</summary>
    private async Task LoadTreeFromIndexAsync(string rootPath)
    {
        if (!File.Exists(DefaultDatabasePath)) return;

        DiskNode rootNode = await Task.Run(() =>
        {
            using var store = new DiskIndexStore(DefaultDatabasePath);
            IReadOnlyList<IndexedFileRecord> rawFiles = store.ListFilesUnder(rootPath);
            var files = rawFiles
                .Where(f => f.FilePath.IndexOf("Recycle", StringComparison.OrdinalIgnoreCase) < 0)
                .ToList();
            return BuildDiskNodeTree(rootPath, files);
        });

        _rootNodes.Clear();
        _rootNodes.Add(rootNode);
        ExpandRootNode();
    }

    private static DiskNode BuildDiskNodeTree(string rootPath, IReadOnlyList<IndexedFileRecord> files)
    {
        string normalizedRoot = Path.GetFullPath(rootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        // Drive roots like "D:" become empty after TrimEnd — pad with separator so display + lookups work.
        if (normalizedRoot.Length == 2 && normalizedRoot[1] == ':')
        {
            normalizedRoot += Path.DirectorySeparatorChar;
        }

        string rootDisplayName = Path.GetFileName(normalizedRoot);
        if (string.IsNullOrEmpty(rootDisplayName)) rootDisplayName = normalizedRoot;

        var rootNode = new DiskNode
        {
            Name        = rootDisplayName,
            FullPath    = normalizedRoot,
            IsDirectory = true,
            SizeBytes   = 0
        };

        // Map directory full-path → node, so successive files reuse parent dir nodes.
        var dirByPath = new Dictionary<string, DiskNode>(StringComparer.OrdinalIgnoreCase)
        {
            [normalizedRoot.TrimEnd(Path.DirectorySeparatorChar)] = rootNode,
            [normalizedRoot] = rootNode
        };

        foreach (IndexedFileRecord file in files)
        {
            string? dirPath = Path.GetDirectoryName(file.FilePath);
            if (string.IsNullOrEmpty(dirPath)) continue;

            DiskNode parent = EnsureDirectoryChain(dirPath, dirByPath, normalizedRoot);
            parent.Children.Add(new DiskNode
            {
                Name        = file.FileName,
                FullPath    = file.FilePath,
                IsDirectory = false,
                SizeBytes   = file.FileSize
            });
        }

        long rootSize = AggregateSizes(rootNode);
        ApplyPercent(rootNode, Math.Max(rootSize, 1));
        SortTreeChildren(rootNode);
        return rootNode;
    }

    private static DiskNode EnsureDirectoryChain(
        string dirPath,
        Dictionary<string, DiskNode> dirByPath,
        string normalizedRoot)
    {
        if (dirByPath.TryGetValue(dirPath, out DiskNode? existing))
        {
            return existing;
        }

        string? parentPath = Path.GetDirectoryName(dirPath);
        DiskNode parent = string.IsNullOrEmpty(parentPath)
            ? dirByPath[normalizedRoot.TrimEnd(Path.DirectorySeparatorChar)]
            : EnsureDirectoryChain(parentPath, dirByPath, normalizedRoot);

        var node = new DiskNode
        {
            Name        = Path.GetFileName(dirPath) ?? dirPath,
            FullPath    = dirPath,
            IsDirectory = true,
            SizeBytes   = 0
        };
        parent.Children.Add(node);
        dirByPath[dirPath] = node;
        return node;
    }

    private static long AggregateSizes(DiskNode node)
    {
        if (!node.IsDirectory) return node.SizeBytes;
        long total = 0;
        foreach (DiskNode child in node.Children) total += AggregateSizes(child);
        node.SizeBytes = total;
        return total;
    }

    private static void ApplyPercent(DiskNode node, long rootSize)
    {
        node.PercentOfRoot = (double)node.SizeBytes / rootSize * 100.0;
        foreach (DiskNode child in node.Children) ApplyPercent(child, rootSize);
    }

    /// <summary>Returns groups of folders whose subtrees are byte-identical AND whose leaf folder
    /// names match. Each group has at least 2 members. Empty folders and single-file folders are
    /// excluded — those are already covered by file-level duplicate detection and would just add noise.
    /// Result tuples align by index: paths[i] ↔ totalSize[i] ↔ lastWrite[i].</summary>
    private static List<(IReadOnlyList<string> Paths, IReadOnlyList<long> TotalSizes, IReadOnlyList<DateTime> LastWrites)>
        ComputeDuplicateFolderGroups(IReadOnlyList<IndexedFileRecord> files)
    {
        if (files.Count == 0) return new();

        // Build a directory-only tree from the flat file list, recording each dir's direct files
        // and child dirs. Directory paths use case-insensitive matching to handle Windows.
        var dirByPath = new Dictionary<string, FolderFingerprintNode>(StringComparer.OrdinalIgnoreCase);

        FolderFingerprintNode GetOrCreate(string path)
        {
            if (dirByPath.TryGetValue(path, out FolderFingerprintNode? existing)) return existing;
            var node = new FolderFingerprintNode { Path = path };
            dirByPath[path] = node;
            return node;
        }

        foreach (IndexedFileRecord f in files)
        {
            FolderFingerprintNode dir = GetOrCreate(f.DirectoryPath);
            dir.DirectFiles.Add(f);
        }

        // Walk each dir up to the filesystem root, lazily creating parent nodes and wiring
        // child→parent edges. Ancestors above the scope can't form duplicate groups (they're
        // singletons in this dir set) so they harmlessly pass through.
        foreach (string dirPath in dirByPath.Keys.ToList())
        {
            string current = dirPath;
            while (true)
            {
                string? parent = Path.GetDirectoryName(current);
                if (string.IsNullOrEmpty(parent)) break;
                if (string.Equals(parent, current, StringComparison.OrdinalIgnoreCase)) break;

                FolderFingerprintNode parentNode = GetOrCreate(parent);
                FolderFingerprintNode currentNode = dirByPath[current];
                if (!parentNode.SubDirs.Any(d => string.Equals(d.Path, current, StringComparison.OrdinalIgnoreCase)))
                {
                    parentNode.SubDirs.Add(currentNode);
                }
                current = parent;
            }
        }

        using var sha = System.Security.Cryptography.SHA256.Create();
        foreach (FolderFingerprintNode node in dirByPath.Values)
        {
            ComputeFingerprint(node, sha);
        }

        // Group by (last-segment name + fingerprint). Skip empty / single-file folders to keep noise down.
        var groupsRaw = dirByPath.Values
            .Where(d => d.DescendantFileCount >= 2 && !string.IsNullOrEmpty(d.Fingerprint))
            .GroupBy(d => $"{(Path.GetFileName(d.Path) ?? string.Empty).ToLowerInvariant()}|{d.Fingerprint}")
            .Where(g => g.Count() > 1)
            .OrderByDescending(g => g.First().TotalSize);

        var result = new List<(IReadOnlyList<string>, IReadOnlyList<long>, IReadOnlyList<DateTime>)>();
        foreach (var group in groupsRaw)
        {
            var members = group
                .OrderBy(d => d.Path, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var paths      = members.Select(m => m.Path).ToList();
            var sizes      = members.Select(m => m.TotalSize).ToList();
            var lastWrites = members.Select(m => SafeDirLastWriteUtc(m.Path)).ToList();
            result.Add((paths, sizes, lastWrites));
        }
        return result;
    }

    private sealed class FolderFingerprintNode
    {
        public string Path = string.Empty;
        public List<IndexedFileRecord> DirectFiles = new();
        public List<FolderFingerprintNode> SubDirs = new();
        public string? Fingerprint;
        public long DescendantFileCount;
        public long TotalSize;
    }

    private static void ComputeFingerprint(FolderFingerprintNode node, System.Security.Cryptography.SHA256 sha)
    {
        if (node.Fingerprint != null) return;

        long fileCount = node.DirectFiles.Count;
        long totalSize = node.DirectFiles.Sum(f => f.FileSize);

        foreach (FolderFingerprintNode sd in node.SubDirs)
        {
            ComputeFingerprint(sd, sha);
            fileCount += sd.DescendantFileCount;
            totalSize += sd.TotalSize;
        }

        // Canonical content string: files first (sorted by name), then subdirs (sorted by leaf name).
        // Files contribute name + size + headTailHash; subdirs contribute name + their fingerprint.
        var sb = new System.Text.StringBuilder();
        foreach (IndexedFileRecord f in node.DirectFiles.OrderBy(x => x.FileName, StringComparer.OrdinalIgnoreCase))
        {
            sb.Append("F:").Append(f.FileName).Append(':').Append(f.FileSize).Append(':').Append(f.HeadTailHash).Append('\n');
        }
        foreach (FolderFingerprintNode sd in node.SubDirs.OrderBy(d => Path.GetFileName(d.Path) ?? string.Empty, StringComparer.OrdinalIgnoreCase))
        {
            sb.Append("D:").Append(Path.GetFileName(sd.Path) ?? string.Empty).Append(':').Append(sd.Fingerprint).Append('\n');
        }

        byte[] hash = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(sb.ToString()));
        node.Fingerprint = Convert.ToHexString(hash);
        node.DescendantFileCount = fileCount;
        node.TotalSize = totalSize;
    }

    private static DateTime SafeDirLastWriteUtc(string path)
    {
        try { return Directory.GetLastWriteTimeUtc(path); }
        catch { return DateTime.MinValue; }
    }

    private static void SortTreeChildren(DiskNode node)
    {
        if (node.Children.Count == 0) return;

        var sorted = node.Children
            .OrderByDescending(c => c.IsDirectory)            // dirs before files
            .ThenByDescending(c => c.SizeBytes)               // bigger first within each kind
            .ThenBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        node.Children.Clear();
        foreach (DiskNode c in sorted) node.Children.Add(c);
        foreach (DiskNode c in node.Children) SortTreeChildren(c);
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

        string tempRoot = TempFolderPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(tempRoot))
        {
            MessageBox.Show("Specify a Temp folder first (top of the window).",
                "Temp Folder Missing", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Snapshot the source + group members before we mutate anything; the move below removes them
        // from disk and the upcoming refresh will clear _duplicateMatches.
        string sourcePath = _selectedHashSourceFilePath;
        var groupMembers = _duplicateMatches.Select(m => m.FilePath).ToList();

        SetBusyState(true);

        try
        {
            StatusTextBlock.Text = "Moving identical files into temp folder...";

            (string targetFolder, int movedCount, int skippedCount, IReadOnlyList<string> movedPaths) =
                await Task.Run(() => MoveCurrentMatchesToTemp(sourcePath, groupMembers, tempRoot));

            // Drop the moved files from the index DB so they stop showing up as duplicates.
            // This is the "이동하고 나면 중복창에서 제외" requirement — without this the DB still
            // points to paths that no longer exist and the middle pane keeps showing stale rows.
            if (movedPaths.Count > 0)
            {
                await Task.Run(() =>
                {
                    using var store = new DiskIndexStore(DefaultDatabasePath);
                    store.DeleteFiles(movedPaths);
                });
            }

            StatusTextBlock.Text = $"Moved {movedCount:N0} files (skipped {skippedCount:N0}) → {targetFolder}";
            DuplicateSummaryTextBlock.Text = $"Moved to: {targetFolder}";

            try
            {
                Process.Start(new ProcessStartInfo { FileName = targetFolder, UseShellExecute = true });
            }
            catch
            {
                // Opening Explorer is best-effort; the move itself already succeeded.
            }

            // Refresh both the tree (folder sizes shifted) and the middle pane (duplicate group is gone).
            string rootPath = RootFolderPathTextBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(rootPath) && Directory.Exists(rootPath))
            {
                await LoadTreeFromIndexAsync(rootPath);
                await LoadDuplicateGroupsInScopeAsync(rootPath);
            }
            else
            {
                _duplicateMatches.Clear();
                _selectedHashSourceFilePath = null;
            }

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

    private static string? LoadLastRootPath()
    {
        try
        {
            if (!File.Exists(LastRootSettingsPath)) return null;
            string text = File.ReadAllText(LastRootSettingsPath).Trim();
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
        catch
        {
            return null;
        }
    }

    private static void SaveLastRootPath(string path)
    {
        try
        {
            string? dir = Path.GetDirectoryName(LastRootSettingsPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(LastRootSettingsPath, path);
        }
        catch
        {
            // Best-effort persistence — UI state is not critical enough to surface a failure.
        }
    }

    private static string? LoadLastTempPath()
    {
        try
        {
            if (!File.Exists(LastTempFolderSettingsPath)) return null;
            string text = File.ReadAllText(LastTempFolderSettingsPath).Trim();
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
        catch
        {
            return null;
        }
    }

    private static void SaveLastTempPath(string path)
    {
        try
        {
            string? dir = Path.GetDirectoryName(LastTempFolderSettingsPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(LastTempFolderSettingsPath, path);
        }
        catch
        {
            // Best-effort persistence — UI state is not critical enough to surface a failure.
        }
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
        ResetPaneSources();
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

    /// <summary>Moves the source file plus all its duplicate copies into a fresh timestamped subfolder
    /// under <paramref name="tempFolderRoot"/>. Returns the actual paths that were successfully moved
    /// so the caller can drop them from the index DB.</summary>
    private static (string TargetFolder, int MovedCount, int SkippedCount, IReadOnlyList<string> MovedPaths)
        MoveCurrentMatchesToTemp(
            string sourceFilePath,
            IReadOnlyList<string> groupMemberPaths,
            string tempFolderRoot)
    {
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string sourceName = Path.GetFileNameWithoutExtension(sourceFilePath);
        string safeSourceName = MakeSafeFileName(string.IsNullOrWhiteSpace(sourceName) ? "selected" : sourceName);

        string targetFolder = Path.Combine(tempFolderRoot, $"{timestamp}_{safeSourceName}");
        Directory.CreateDirectory(targetFolder);

        // Source first so it gets the 0001_ prefix; OrdinalIgnoreCase distinct dedupes if the source
        // is also listed in the group members.
        var orderedPaths = new List<string> { sourceFilePath };
        orderedPaths.AddRange(groupMemberPaths);
        List<string> uniquePaths = orderedPaths
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var mapLines = new List<string>
        {
            $"MovedAt\t{DateTime.Now:O}",
            "MovedName\tOriginalPath"
        };

        int movedCount = 0;
        int skippedCount = 0;
        var movedPaths = new List<string>(uniquePaths.Count);

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
                // File.Move handles cross-volume moves transparently in .NET (falls back to copy+delete).
                File.Move(sourcePath, targetPath);
                mapLines.Add($"{targetFileName}\t{sourcePath}");
                movedCount++;
                movedPaths.Add(sourcePath);
            }
            catch
            {
                skippedCount++;
            }
        }

        try
        {
            File.WriteAllLines(Path.Combine(targetFolder, "_source_map.tsv"), mapLines);
        }
        catch
        {
            // Manifest is a nice-to-have; failure here shouldn't undo a successful move.
        }

        return (targetFolder, movedCount, skippedCount, movedPaths);
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
