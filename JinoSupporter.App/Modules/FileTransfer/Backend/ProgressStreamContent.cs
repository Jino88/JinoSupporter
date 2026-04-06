using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net;

namespace QuickShareClone.Server;

internal sealed class ProgressStreamContent(
    Stream stream,
    long totalBytes,
    Action<long> onProgress,
    Action? onTransferStarted = null) : HttpContent
{
    protected override async Task SerializeToStreamAsync(Stream targetStream, TransportContext? context)
    {
        var buffer = new byte[81920];
        long totalSent = 0;
        int bytesRead;
        var started = false;
        while ((bytesRead = await stream.ReadAsync(buffer)) > 0)
        {
            if (!started)
            {
                started = true;
                onTransferStarted?.Invoke();
            }

            await targetStream.WriteAsync(buffer.AsMemory(0, bytesRead));
            totalSent += bytesRead;
            onProgress(totalSent);
        }
    }

    protected override bool TryComputeLength(out long length)
    {
        length = totalBytes;
        return totalBytes >= 0;
    }
}
