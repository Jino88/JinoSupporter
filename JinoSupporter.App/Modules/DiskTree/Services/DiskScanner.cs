using DiskTree.Models;
using System.IO;

namespace DiskTree.Services;

public readonly record struct ScanProgress(int ProcessedFiles, int TotalFiles, string StatusText);

public sealed class DiskScanner
{
    public int EstimateFileCount(string rootPath, CancellationToken cancellationToken, IProgress<string>? statusProgress = null)
    {
        if (string.IsNullOrWhiteSpace(rootPath))
        {
            throw new ArgumentException("Root path is required.", nameof(rootPath));
        }

        string normalizedRootPath = Path.GetFullPath(rootPath);
        if (!Directory.Exists(normalizedRootPath))
        {
            throw new DirectoryNotFoundException($"Directory not found: {normalizedRootPath}");
        }

        var counters = new EstimateCounters();
        int totalFiles = EstimateDirectoryFileCount(normalizedRootPath, counters, cancellationToken, statusProgress);
        statusProgress?.Report($"Estimated files: {totalFiles:N0}");
        return totalFiles;
    }

    public ScanResult Scan(
        string rootPath,
        int estimatedFileCount,
        CancellationToken cancellationToken,
        IReadOnlyDictionary<string, IndexedFileSnapshot>? existingIndex = null,
        IProgress<ScanProgress>? progress = null)
    {
        if (string.IsNullOrWhiteSpace(rootPath))
        {
            throw new ArgumentException("Root path is required.", nameof(rootPath));
        }

        string normalizedRootPath = Path.GetFullPath(rootPath);
        if (!Directory.Exists(normalizedRootPath))
        {
            throw new DirectoryNotFoundException($"Directory not found: {normalizedRootPath}");
        }

        var indexedFiles = new List<IndexedFileRecord>(capacity: 4096);
        var counters = new ScanCounters();

        DiskNode rootNode = ScanDirectory(
            normalizedRootPath,
            indexedFiles,
            counters,
            estimatedFileCount,
            cancellationToken,
            existingIndex,
            progress);
        ApplyPercent(rootNode, rootNode.SizeBytes);

        progress?.Report(new ScanProgress(
            counters.ScannedFileCount,
            estimatedFileCount,
            $"Scanning done: {counters.ScannedFileCount:N0}/{Math.Max(estimatedFileCount, counters.ScannedFileCount):N0} files | rehash {counters.RehashedCount:N0}, reused {counters.ReusedHashCount:N0}"));

        return new ScanResult
        {
            RootNode = rootNode,
            IndexedFiles = indexedFiles,
            ScannedDirectoryCount = counters.ScannedDirectoryCount,
            ScannedFileCount = counters.ScannedFileCount,
            SkippedDirectoryCount = counters.SkippedDirectoryCount,
            SkippedFileCount = counters.SkippedFileCount,
            RehashedCount = counters.RehashedCount,
            ReusedHashCount = counters.ReusedHashCount
        };
    }

    private static DiskNode ScanDirectory(
        string directoryPath,
        List<IndexedFileRecord> indexedFiles,
        ScanCounters counters,
        int estimatedFileCount,
        CancellationToken cancellationToken,
        IReadOnlyDictionary<string, IndexedFileSnapshot>? existingIndex,
        IProgress<ScanProgress>? progress)
    {
        cancellationToken.ThrowIfCancellationRequested();

        counters.ScannedDirectoryCount++;
        var node = new DiskNode
        {
            Name = GetDirectoryDisplayName(directoryPath),
            FullPath = directoryPath,
            IsDirectory = true,
            SizeBytes = 0
        };

        var childNodes = new List<DiskNode>();

        foreach (string childDirectoryPath in GetDirectoriesSafe(directoryPath, counters))
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (ShouldSkipDirectory(childDirectoryPath))
            {
                counters.SkippedDirectoryCount++;
                continue;
            }

            DiskNode childDirectoryNode = ScanDirectory(
                childDirectoryPath,
                indexedFiles,
                counters,
                estimatedFileCount,
                cancellationToken,
                existingIndex,
                progress);

            childNodes.Add(childDirectoryNode);
            node.SizeBytes += childDirectoryNode.SizeBytes;
        }

        foreach (string filePath in GetFilesSafe(directoryPath, counters))
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!TryBuildFileNode(filePath, existingIndex, out DiskNode? fileNode, out IndexedFileRecord? record, out bool reusedHash))
            {
                counters.SkippedFileCount++;
                continue;
            }

            if (reusedHash)
            {
                counters.ReusedHashCount++;
            }
            else
            {
                counters.RehashedCount++;
            }

            childNodes.Add(fileNode);
            indexedFiles.Add(record);
            node.SizeBytes += fileNode.SizeBytes;
            counters.ScannedFileCount++;

            if (counters.ScannedFileCount % 25 == 0 || counters.ScannedFileCount == estimatedFileCount)
            {
                progress?.Report(new ScanProgress(
                    counters.ScannedFileCount,
                    estimatedFileCount,
                    $"Scanning: {counters.ScannedFileCount:N0}/{Math.Max(estimatedFileCount, counters.ScannedFileCount):N0} | rehash {counters.RehashedCount:N0}, reused {counters.ReusedHashCount:N0}"));
            }
        }

        foreach (DiskNode child in childNodes
            .OrderByDescending(item => item.IsDirectory)
            .ThenByDescending(item => item.SizeBytes)
            .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
        {
            node.Children.Add(child);
        }

        return node;
    }

    private static void ApplyPercent(DiskNode node, long rootSizeBytes)
    {
        node.PercentOfRoot = rootSizeBytes <= 0
            ? 0.0
            : node.SizeBytes * 100.0 / rootSizeBytes;

        foreach (DiskNode child in node.Children)
        {
            ApplyPercent(child, rootSizeBytes);
        }
    }

    private static bool ShouldSkipDirectory(string directoryPath)
    {
        try
        {
            FileAttributes attributes = File.GetAttributes(directoryPath);
            return attributes.HasFlag(FileAttributes.ReparsePoint);
        }
        catch
        {
            return true;
        }
    }

    private static bool TryBuildFileNode(
        string filePath,
        IReadOnlyDictionary<string, IndexedFileSnapshot>? existingIndex,
        out DiskNode fileNode,
        out IndexedFileRecord fileRecord,
        out bool reusedHash)
    {
        fileNode = null!;
        fileRecord = null!;
        reusedHash = false;

        try
        {
            var fileInfo = new FileInfo(filePath);
            if (!fileInfo.Exists)
            {
                return false;
            }

            long fileSize = fileInfo.Length;
            string fullPath = fileInfo.FullName;
            DateTime lastWriteUtc = fileInfo.LastWriteTimeUtc;

            string headTailHash;
            if (existingIndex is not null &&
                existingIndex.TryGetValue(fullPath, out IndexedFileSnapshot snapshot) &&
                snapshot.FileSize == fileSize &&
                snapshot.LastWriteUtc == lastWriteUtc)
            {
                headTailHash = snapshot.HeadTailHash;
                reusedHash = true;
            }
            else
            {
                headTailHash = FileHasher.ComputeHeadTailHash(fullPath, fileSize);
            }

            DateTime scannedUtc = DateTime.UtcNow;

            fileNode = new DiskNode
            {
                Name = fileInfo.Name,
                FullPath = fullPath,
                IsDirectory = false,
                SizeBytes = fileSize
            };

            fileRecord = new IndexedFileRecord(
                fullPath,
                fileInfo.Name,
                fileInfo.DirectoryName ?? string.Empty,
                fileSize,
                headTailHash,
                lastWriteUtc,
                scannedUtc);

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string[] GetDirectoriesSafe(string directoryPath, ScanCounters counters)
    {
        try
        {
            return Directory.GetDirectories(directoryPath);
        }
        catch
        {
            counters.SkippedDirectoryCount++;
            return [];
        }
    }

    private static string[] GetFilesSafe(string directoryPath, ScanCounters counters)
    {
        try
        {
            return Directory.GetFiles(directoryPath);
        }
        catch
        {
            counters.SkippedFileCount++;
            return [];
        }
    }

    private static string GetDirectoryDisplayName(string directoryPath)
    {
        string? name = Path.GetFileName(directoryPath);
        return string.IsNullOrWhiteSpace(name) ? directoryPath : name;
    }

    private static int EstimateDirectoryFileCount(
        string directoryPath,
        EstimateCounters counters,
        CancellationToken cancellationToken,
        IProgress<string>? statusProgress)
    {
        cancellationToken.ThrowIfCancellationRequested();
        counters.VisitedDirectoryCount++;

        int totalFiles = 0;

        foreach (string childDirectoryPath in GetDirectoriesSafe(directoryPath))
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (ShouldSkipDirectory(childDirectoryPath))
            {
                counters.SkippedDirectoryCount++;
                continue;
            }

            totalFiles += EstimateDirectoryFileCount(childDirectoryPath, counters, cancellationToken, statusProgress);
        }

        foreach (string filePath in GetFilesSafe(directoryPath))
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Exists)
                {
                    totalFiles++;
                }
            }
            catch
            {
                counters.SkippedFileCount++;
            }
        }

        if (counters.VisitedDirectoryCount % 300 == 0)
        {
            statusProgress?.Report($"Estimating files... visited {counters.VisitedDirectoryCount:N0} folders");
        }

        return totalFiles;
    }

    private static string[] GetDirectoriesSafe(string directoryPath)
    {
        try
        {
            return Directory.GetDirectories(directoryPath);
        }
        catch
        {
            return [];
        }
    }

    private static string[] GetFilesSafe(string directoryPath)
    {
        try
        {
            return Directory.GetFiles(directoryPath);
        }
        catch
        {
            return [];
        }
    }

    private sealed class ScanCounters
    {
        public int ScannedDirectoryCount { get; set; }
        public int ScannedFileCount { get; set; }
        public int SkippedDirectoryCount { get; set; }
        public int SkippedFileCount { get; set; }
        public int ReusedHashCount { get; set; }
        public int RehashedCount { get; set; }
    }

    private sealed class EstimateCounters
    {
        public int VisitedDirectoryCount { get; set; }
        public int SkippedDirectoryCount { get; set; }
        public int SkippedFileCount { get; set; }
    }
}
