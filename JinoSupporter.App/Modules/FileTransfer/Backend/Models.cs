namespace QuickShareClone.Server;

public sealed record UploadStatusResponse(
    string FileId,
    IReadOnlyCollection<int> ReceivedChunks,
    bool DestinationSelected);

public sealed record UploadRequest(string FileId, string FileName, int TotalChunks, long TotalBytes);

public sealed record UploadDestinationRequest(string FileId, string DestinationDirectory);

public sealed record CompleteUploadRequest(string FileId, string FileName, int TotalChunks);

public sealed record UploadSessionSummary(
    string FileId,
    string FileName,
    int ReceivedChunkCount,
    int? TotalChunks,
    long ReceivedBytes,
    long? TotalBytes,
    bool DestinationSelected,
    string? DestinationDirectory,
    string? FinalFilePath,
    bool IsCompleted,
    DateTimeOffset StartedAt,
    DateTimeOffset UpdatedAt);

public sealed record DiscoveryDeviceInfo(
    string DeviceId,
    string DeviceName,
    string Platform,
    IReadOnlyCollection<string> ServerUrls,
    DateTimeOffset LastUpdatedAt);

public sealed record AndroidDeviceRegistrationRequest(
    string DeviceId,
    string DeviceName,
    string Platform,
    string ReceiveUrl);

public sealed record AndroidTransferResult(
    string FileName,
    bool Success,
    string Message);

public sealed record AndroidTransferOfferRequest(
    string OfferId,
    int FileCount,
    long TotalBytes,
    IReadOnlyCollection<AndroidTransferOfferFile> Files);

public sealed record AndroidBrowserSendOfferRequest(
    string DeviceId,
    IReadOnlyCollection<AndroidTransferOfferFile> Files);

public sealed record AndroidTransferOfferFile(
    string FileName,
    long FileSizeBytes);

public sealed record AndroidConnectedDevice(
    string DeviceId,
    string DeviceName,
    string Platform,
    string ReceiveUrl,
    DateTimeOffset LastSeenAt);

public sealed record AndroidOutboundTransferSummary(
    string TransferId,
    string DeviceId,
    string DeviceName,
    string FileName,
    long TotalBytes,
    long SentBytes,
    bool IsCompleted,
    string StatusText,
    DateTimeOffset StartedAt,
    DateTimeOffset UpdatedAt);

public sealed class UploadSession
{
    public required string FileId { get; init; }
    public required string FileName { get; set; }
    public int? TotalChunks { get; set; }
    public long? TotalBytes { get; set; }
    public long ReceivedBytes { get; set; }
    public string? DestinationDirectory { get; set; }
    public string? FinalFilePath { get; set; }
    public bool IsCompleted { get; set; }
    public HashSet<int> ReceivedChunks { get; } = [];
    public DateTimeOffset StartedAt { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset UpdatedAt { get; set; } = DateTimeOffset.UtcNow;
}

public sealed class AndroidOutboundTransfer
{
    public required string TransferId { get; init; }
    public required string DeviceId { get; set; }
    public required string DeviceName { get; set; }
    public required string FileName { get; set; }
    public long TotalBytes { get; set; }
    public long SentBytes { get; set; }
    public bool IsCompleted { get; set; }
    public string StatusText { get; set; } = "Preparing";
    public DateTimeOffset StartedAt { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset UpdatedAt { get; set; } = DateTimeOffset.UtcNow;
}
