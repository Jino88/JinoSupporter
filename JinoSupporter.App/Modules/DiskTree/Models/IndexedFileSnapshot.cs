namespace DiskTree.Models;

public readonly record struct IndexedFileSnapshot(
    long FileSize,
    DateTime LastWriteUtc,
    string HeadTailHash);
