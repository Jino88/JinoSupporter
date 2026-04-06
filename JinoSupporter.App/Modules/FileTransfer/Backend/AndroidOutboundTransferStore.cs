using System.Collections.Concurrent;

namespace QuickShareClone.Server;

public sealed class AndroidOutboundTransferStore
{
    private readonly ConcurrentDictionary<string, AndroidOutboundTransfer> _transfers = new();
    private readonly ConcurrentDictionary<string, List<string>> _offerTransfers = new();

    public AndroidOutboundTransfer Start(string deviceId, string deviceName, string fileName, long totalBytes)
    {
        var transfer = new AndroidOutboundTransfer
        {
            TransferId = Guid.NewGuid().ToString("N"),
            DeviceId = deviceId,
            DeviceName = deviceName,
            FileName = fileName,
            TotalBytes = totalBytes,
            StatusText = "Waiting for Android approval"
        };

        _transfers[transfer.TransferId] = transfer;
        return transfer;
    }

    public void RegisterPendingOffer(string offerId, string deviceId, string deviceName, IReadOnlyCollection<AndroidTransferOfferFile> files, string statusText)
    {
        var transferIds = new List<string>(files.Count);
        foreach (var file in files)
        {
            var transfer = new AndroidOutboundTransfer
            {
                TransferId = Guid.NewGuid().ToString("N"),
                DeviceId = deviceId,
                DeviceName = deviceName,
                FileName = file.FileName,
                TotalBytes = file.FileSizeBytes,
                StatusText = statusText
            };

            _transfers[transfer.TransferId] = transfer;
            transferIds.Add(transfer.TransferId);
        }

        _offerTransfers[offerId] = transferIds;
    }

    public AndroidOutboundTransfer GetOrAttachToOffer(string offerId, string deviceId, string deviceName, string fileName, long totalBytes)
    {
        if (_offerTransfers.TryGetValue(offerId, out var transferIds))
        {
            foreach (var transferId in transferIds)
            {
                if (!_transfers.TryGetValue(transferId, out var transfer))
                {
                    continue;
                }

                if (!string.Equals(transfer.FileName, fileName, StringComparison.Ordinal))
                {
                    continue;
                }

                lock (transfer)
                {
                    transfer.DeviceId = deviceId;
                    transfer.DeviceName = deviceName;
                    transfer.TotalBytes = totalBytes;
                    transfer.StatusText = "Sending to Android";
                    transfer.UpdatedAt = DateTimeOffset.UtcNow;
                }

                return transfer;
            }
        }

        return Start(deviceId, deviceName, fileName, totalBytes);
    }

    public void UpdateOfferStatus(string offerId, string statusText, bool isCompleted = false)
    {
        if (!_offerTransfers.TryGetValue(offerId, out var transferIds))
        {
            return;
        }

        foreach (var transferId in transferIds)
        {
            if (!_transfers.TryGetValue(transferId, out var transfer))
            {
                continue;
            }

            lock (transfer)
            {
                transfer.StatusText = statusText;
                transfer.IsCompleted = isCompleted;
                transfer.UpdatedAt = DateTimeOffset.UtcNow;
            }
        }
    }

    public void UpdateProgress(string transferId, long sentBytes, string? statusText = null)
    {
        if (!_transfers.TryGetValue(transferId, out var transfer))
        {
            return;
        }

        lock (transfer)
        {
            transfer.SentBytes = Math.Max(transfer.SentBytes, sentBytes);
            if (!string.IsNullOrWhiteSpace(statusText))
            {
                transfer.StatusText = statusText;
            }

            transfer.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void Complete(string transferId, string statusText)
    {
        if (!_transfers.TryGetValue(transferId, out var transfer))
        {
            return;
        }

        lock (transfer)
        {
            transfer.SentBytes = transfer.TotalBytes;
            transfer.IsCompleted = true;
            transfer.StatusText = statusText;
            transfer.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public void Fail(string transferId, string statusText)
    {
        if (!_transfers.TryGetValue(transferId, out var transfer))
        {
            return;
        }

        lock (transfer)
        {
            transfer.IsCompleted = true;
            transfer.StatusText = statusText;
            transfer.UpdatedAt = DateTimeOffset.UtcNow;
        }
    }

    public IReadOnlyCollection<AndroidOutboundTransferSummary> List() =>
        _transfers.Values
            .OrderByDescending(x => x.UpdatedAt)
            .Select(x => new AndroidOutboundTransferSummary(
                x.TransferId,
                x.DeviceId,
                x.DeviceName,
                x.FileName,
                x.TotalBytes,
                x.SentBytes,
                x.IsCompleted,
                x.StatusText,
                x.StartedAt,
                x.UpdatedAt))
            .ToArray();
}
