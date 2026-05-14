using DiskTree.Services;

namespace DiskTree.Models;

public sealed record DuplicateMatchRecord(
    string FilePath,
    long FileSize,
    DateTime LastWriteUtc,
    int GroupId = 0,
    int GroupSize = 0,
    bool IsDirectory = false)
{
    public string SizeText      => ByteSizeFormatter.Format(FileSize);
    public string LastWriteText => LastWriteUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
    public string GroupLabel    => GroupSize > 0
        ? (IsDirectory ? $"DIR #{GroupId} ({GroupSize})" : $"#{GroupId} ({GroupSize})")
        : string.Empty;
    // Folder rows get a leading marker + trailing slash so they're visually distinct from file rows.
    public string PathDisplay   => IsDirectory ? $"[FOLDER] {FilePath}\\" : FilePath;
}
