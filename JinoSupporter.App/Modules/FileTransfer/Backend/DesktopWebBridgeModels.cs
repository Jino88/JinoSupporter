namespace QuickShareClone.Server;

public sealed record DesktopNativeSelectionRequest(
    IReadOnlyCollection<DesktopSelectedFile> Files);

public sealed record DesktopSelectedFile(
    string FileName,
    string FilePath,
    long FileSizeBytes);

public sealed record DesktopNativeSelection(
    string SelectionId,
    IReadOnlyCollection<DesktopSelectedFile> Files,
    DateTimeOffset CreatedAt);

public sealed record DesktopNativeSendRequest(
    string DeviceId,
    string SelectionId,
    string OfferId);
