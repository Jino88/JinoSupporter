using System.Collections.Concurrent;

namespace QuickShareClone.Server;

public sealed class UploadStore
{
    private readonly ConcurrentDictionary<string, UploadSession> _sessions = new();

    public UploadSession GetOrCreate(string fileId, string fileName, int? totalChunks = null, long? totalBytes = null)
    {
        var session = _sessions.GetOrAdd(fileId, _ => new UploadSession
        {
            FileId = fileId,
            FileName = fileName,
            StartedAt = DateTimeOffset.UtcNow
        });

        lock (session)
        {
            session.FileName = fileName;
            if (totalChunks.HasValue)
            {
                session.TotalChunks = Math.Max(session.TotalChunks ?? 0, totalChunks.Value);
            }

            if (totalBytes.HasValue)
            {
                session.TotalBytes = Math.Max(session.TotalBytes ?? 0, totalBytes.Value);
            }

            session.UpdatedAt = DateTimeOffset.UtcNow;
        }

        return session;
    }

    public UploadSession? Find(string fileId) => _sessions.TryGetValue(fileId, out var session) ? session : null;

    public IReadOnlyCollection<UploadSessionSummary> List() =>
        _sessions.Values
            .OrderByDescending(x => x.UpdatedAt)
            .Select(x => new UploadSessionSummary(
                x.FileId,
                x.FileName,
                x.ReceivedChunks.Count,
                x.TotalChunks,
                x.ReceivedBytes,
                x.TotalBytes,
                !string.IsNullOrWhiteSpace(x.DestinationDirectory),
                x.DestinationDirectory,
                x.FinalFilePath,
                x.IsCompleted,
                x.StartedAt,
                x.UpdatedAt))
            .ToArray();

    public void MarkReceived(string fileId, int chunkIndex, long bytesReceived)
    {
        if (!_sessions.TryGetValue(fileId, out var session))
        {
            return;
        }

        lock (session)
        {
            session.ReceivedChunks.Add(chunkIndex);
            session.ReceivedBytes += bytesReceived;
            session.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void MarkCompleted(string fileId)
    {
        if (!_sessions.TryGetValue(fileId, out var session))
        {
            return;
        }

        lock (session)
        {
            session.IsCompleted = true;
            session.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void SetDestination(string fileId, string destinationDirectory)
    {
        if (!_sessions.TryGetValue(fileId, out var session))
        {
            return;
        }

        lock (session)
        {
            session.DestinationDirectory = destinationDirectory;
            session.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void SetFinalPath(string fileId, string finalFilePath)
    {
        if (!_sessions.TryGetValue(fileId, out var session))
        {
            return;
        }

        lock (session)
        {
            session.FinalFilePath = finalFilePath;
            session.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }
}
