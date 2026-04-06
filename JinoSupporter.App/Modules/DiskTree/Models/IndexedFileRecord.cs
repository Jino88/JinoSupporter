namespace DiskTree.Models;

public sealed record IndexedFileRecord(
    string FilePath,
    string FileName,
    string DirectoryPath,
    long FileSize,
    string HeadTailHash,
    DateTime LastWriteUtc,
    DateTime ScannedUtc);
