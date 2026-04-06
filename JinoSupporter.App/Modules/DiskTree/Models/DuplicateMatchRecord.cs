using DiskTree.Services;

namespace DiskTree.Models;

public sealed record DuplicateMatchRecord(
    string FilePath,
    long FileSize,
    DateTime LastWriteUtc)
{
    public string SizeText => ByteSizeFormatter.Format(FileSize);
    public string LastWriteText => LastWriteUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
}
