namespace DiskTree.Models;

public sealed class ScanResult
{
    public required DiskNode RootNode { get; init; }
    public required IReadOnlyList<IndexedFileRecord> IndexedFiles { get; init; }
    public int ScannedDirectoryCount { get; init; }
    public int ScannedFileCount { get; init; }
    public int SkippedDirectoryCount { get; init; }
    public int SkippedFileCount { get; init; }
    public int ReusedHashCount { get; init; }
    public int RehashedCount { get; init; }
}
