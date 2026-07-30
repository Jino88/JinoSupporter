using Microsoft.JSInterop;

namespace BmesNgRateStandalone.Services;

public static class BrowserDownload
{
    public const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public static async ValueTask DownloadBytesAsync(
        IJSRuntime js,
        string filename,
        byte[] bytes,
        string contentType)
    {
        try
        {
            await using var stream = new MemoryStream(bytes);
            using var streamRef = new DotNetStreamReference(stream);
            await js.InvokeVoidAsync("downloadFileFromStream", filename, streamRef, contentType);
        }
        catch (JSException)
        {
            string base64 = Convert.ToBase64String(bytes);
            await js.InvokeVoidAsync("downloadBase64File", filename, base64, contentType);
        }
    }
}
