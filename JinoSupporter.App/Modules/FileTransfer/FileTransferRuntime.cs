using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using QuickShareClone.Server;

namespace JinoSupporter.App.Modules.FileTransfer;

internal sealed class FileTransferRuntime : IAsyncDisposable
{
    private static readonly Lazy<FileTransferRuntime> LazyInstance = new(() => new FileTransferRuntime());

    private readonly JsonSerializerOptions _jsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    private readonly Task _startupTask;
    private WebApplication? _app;

    private FileTransferRuntime()
    {
        _startupTask = StartAsync();
    }

    public static FileTransferRuntime Instance => LazyInstance.Value;
    public static bool IsCreated => LazyInstance.IsValueCreated;

    public UploadStore UploadStore => GetRequiredService<UploadStore>();
    public AndroidDeviceStore AndroidDeviceStore => GetRequiredService<AndroidDeviceStore>();
    public AndroidOutboundTransferStore AndroidOutboundTransferStore => GetRequiredService<AndroidOutboundTransferStore>();
    public DeviceIdentityService DeviceIdentityService => GetRequiredService<DeviceIdentityService>();

    public Task EnsureStartedAsync() => _startupTask;

    public async Task<IReadOnlyCollection<AndroidTransferResult>> SendFilesToAndroidAsync(
        string deviceId,
        IReadOnlyCollection<string> filePaths,
        CancellationToken cancellationToken)
    {
        await EnsureStartedAsync();

        AndroidConnectedDevice? device = AndroidDeviceStore.Find(deviceId);
        if (device is null)
        {
            throw new InvalidOperationException("Selected Android device is not registered.");
        }

        if (filePaths.Count == 0)
        {
            throw new InvalidOperationException("Choose at least one file first.");
        }

        var files = filePaths
            .Select(path =>
            {
                FileInfo info = new(path);
                return new AndroidTransferOfferFile(info.Name, info.Length);
            })
            .ToArray();

        string offerId = Guid.NewGuid().ToString("N");
        AndroidOutboundTransferStore.RegisterPendingOffer(offerId, device.DeviceId, device.DeviceName, files, "Waiting for Android approval");

        using HttpClient client = new() { Timeout = TimeSpan.FromMinutes(15) };
        AndroidTransferOfferRequest offer = new(offerId, files.Length, files.Sum(static file => file.FileSizeBytes), files);
        using (HttpRequestMessage offerRequest = new(HttpMethod.Post, $"{device.ReceiveUrl}/api/device/offer"))
        {
            offerRequest.Content = new StringContent(JsonSerializer.Serialize(offer, _jsonOptions), Encoding.UTF8, "application/json");
            using HttpResponseMessage offerResponse = await client.SendAsync(offerRequest, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            string offerBody = await offerResponse.Content.ReadAsStringAsync(cancellationToken);
            if (!offerResponse.IsSuccessStatusCode)
            {
                AndroidOutboundTransferStore.UpdateOfferStatus(offerId, string.IsNullOrWhiteSpace(offerBody) ? "Android rejected transfer." : offerBody, true);
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(offerBody) ? "Android rejected transfer." : offerBody);
            }
        }

        DateTimeOffset approvalDeadline = DateTimeOffset.UtcNow.AddSeconds(60);
        while (DateTimeOffset.UtcNow < approvalDeadline)
        {
            using HttpResponseMessage statusResponse = await client.GetAsync($"{device.ReceiveUrl}/api/device/offer/status?offerId={offerId}", cancellationToken);
            string statusText = await statusResponse.Content.ReadAsStringAsync(cancellationToken);
            if (statusResponse.StatusCode == HttpStatusCode.Conflict)
            {
                AndroidOutboundTransferStore.UpdateOfferStatus(offerId, "Declined on Android", true);
                throw new InvalidOperationException("Transfer declined on Android.");
            }

            if (statusResponse.IsSuccessStatusCode && statusText.Contains("\"approved\"", StringComparison.OrdinalIgnoreCase))
            {
                AndroidOutboundTransferStore.UpdateOfferStatus(offerId, "Approved on Android. Sending files");
                break;
            }

            await Task.Delay(200, cancellationToken);
        }

        List<AndroidTransferResult> results = new(filePaths.Count);
        foreach (string filePath in filePaths)
        {
            FileInfo info = new(filePath);
            AndroidOutboundTransfer transfer = AndroidOutboundTransferStore.GetOrAttachToOffer(offerId, device.DeviceId, device.DeviceName, info.Name, info.Length);
            try
            {
                await using FileStream stream = info.OpenRead();
                using ProgressStreamContent content = new(
                    stream,
                    info.Length,
                    bytesSent => AndroidOutboundTransferStore.UpdateProgress(transfer.TransferId, bytesSent, "Sending to Android"));
                content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");

                using HttpRequestMessage request = new(HttpMethod.Post, $"{device.ReceiveUrl}/api/device/receive");
                request.Content = content;
                request.Headers.Add("X-QuickShare-File-Name-Base64", Convert.ToBase64String(Encoding.UTF8.GetBytes(info.Name)));
                request.Headers.Add("X-QuickShare-Device-Id", device.DeviceId);
                request.Headers.Add("X-QuickShare-Offer-Id", offerId);

                using HttpResponseMessage response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
                string body = await response.Content.ReadAsStringAsync(cancellationToken);
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(body) ? "Android receive failed." : body);
                }

                AndroidOutboundTransferStore.Complete(transfer.TransferId, "Saved on Android");
                results.Add(new AndroidTransferResult(info.Name, true, "Saved on Android"));
            }
            catch (Exception ex)
            {
                AndroidOutboundTransferStore.Fail(transfer.TransferId, ex.Message);
                results.Add(new AndroidTransferResult(info.Name, false, ex.Message));
            }
        }

        return results;
    }

    private async Task StartAsync()
    {
        WebApplicationBuilder builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            ApplicationName = typeof(FileTransferRuntime).Assembly.FullName,
            ContentRootPath = AppContext.BaseDirectory
        });

        builder.WebHost.UseUrls("http://10.6.4.54:5070");
        builder.Services.Configure<FormOptions>(options => options.MultipartBodyLengthLimit = long.MaxValue);
        builder.Services.AddSingleton(Microsoft.Extensions.Options.Options.Create(new UploadOptions()));
        builder.Services.AddSingleton(Microsoft.Extensions.Options.Options.Create(new DiscoveryOptions()));
        builder.Services.AddSingleton<UploadStore>();
        builder.Services.AddSingleton<ChunkFileService>();
        builder.Services.AddSingleton<AndroidDeviceStore>();
        builder.Services.AddSingleton<AndroidOutboundTransferStore>();
        builder.Services.AddSingleton<DesktopFileSelectionStore>();
        builder.Services.AddSingleton<DeviceIdentityService>();
        builder.Services.AddHostedService<CleanupService>();
        builder.Services.AddHostedService<DiscoveryBroadcastService>();

        _app = builder.Build();
        MapEndpoints(_app);
        await _app.StartAsync();
    }

    private void MapEndpoints(WebApplication app)
    {
        app.MapGet("/api/android/devices", (AndroidDeviceStore store) => Results.Ok(store.List()));
        app.MapGet("/api/android/outbound-transfers", (AndroidOutboundTransferStore store) => Results.Ok(store.List()));
        app.MapGet("/api/upload/sessions", (UploadStore store) => Results.Ok(store.List()));
        app.MapGet("/api/device/self", (DeviceIdentityService identityService) => Results.Ok(identityService.GetCurrentDevice()));

        app.MapPost("/api/android/register", (AndroidDeviceRegistrationRequest request, AndroidDeviceStore store) =>
        {
            if (string.IsNullOrWhiteSpace(request.DeviceId) ||
                string.IsNullOrWhiteSpace(request.DeviceName) ||
                string.IsNullOrWhiteSpace(request.ReceiveUrl))
            {
                return Results.BadRequest(new { message = "Invalid Android registration payload." });
            }

            store.Register(request);
            return Results.Ok();
        });

        app.MapPost("/api/upload/request", (UploadRequest request, UploadStore store) =>
        {
            store.GetOrCreate(request.FileId, request.FileName, request.TotalChunks, request.TotalBytes);
            return Results.Accepted();
        });

        app.MapPost("/api/upload/destination", (UploadDestinationRequest request, UploadStore store) =>
        {
            store.SetDestination(request.FileId, request.DestinationDirectory);
            return Results.Ok();
        });

        app.MapPost("/api/upload/chunk", async (HttpRequest request, UploadStore store, ChunkFileService chunks, CancellationToken cancellationToken) =>
        {
            if (!request.HasFormContentType)
            {
                return Results.BadRequest(new { message = "Multipart form-data is required." });
            }

            IFormCollection form = await request.ReadFormAsync(cancellationToken);
            string fileId = form["fileId"].ToString();
            string fileName = form["fileName"].ToString();
            bool chunkOk = int.TryParse(form["chunkIndex"], out int chunkIndex);
            bool totalChunkOk = int.TryParse(form["totalChunks"], out int totalChunks);
            IFormFile? file = form.Files.GetFile("chunk");
            if (!chunkOk || !totalChunkOk || string.IsNullOrWhiteSpace(fileId) || string.IsNullOrWhiteSpace(fileName) || file is null)
            {
                return Results.BadRequest(new { message = "Missing required chunk fields." });
            }

            UploadSession session = store.GetOrCreate(fileId, fileName, totalChunks);
            if (string.IsNullOrWhiteSpace(session.DestinationDirectory))
            {
                return Results.Conflict(new { message = "Destination folder has not been selected on the PC yet." });
            }

            await using Stream stream = file.OpenReadStream();
            long bytesSaved = await chunks.SaveChunkAsync(fileId, chunkIndex, stream, cancellationToken);
            store.MarkReceived(fileId, chunkIndex, bytesSaved);
            return Results.Ok(new { fileId, chunkIndex, totalChunks, saved = true });
        });

        app.MapPost("/api/upload/complete", async (CompleteUploadRequest request, UploadStore store, ChunkFileService chunks, CancellationToken cancellationToken) =>
        {
            IReadOnlyCollection<int> received = chunks.GetReceivedChunks(request.FileId);
            int[] missing = Enumerable.Range(0, request.TotalChunks).Except(received).ToArray();
            if (missing.Length > 0)
            {
                return Results.BadRequest(new { message = "Upload is incomplete.", missingChunks = missing });
            }

            UploadSession session = store.GetOrCreate(request.FileId, request.FileName, request.TotalChunks);
            if (string.IsNullOrWhiteSpace(session.DestinationDirectory))
            {
                return Results.Conflict(new { message = "Destination folder has not been selected on the PC yet." });
            }

            string outputPath = await chunks.MergeChunksAsync(request.FileId, request.FileName, request.TotalChunks, session.DestinationDirectory, cancellationToken);
            store.SetFinalPath(request.FileId, outputPath);
            store.MarkCompleted(request.FileId);
            return Results.Ok(new { request.FileId, outputPath });
        });
    }

    private T GetRequiredService<T>() where T : notnull
    {
        if (_app is null)
        {
            throw new InvalidOperationException("File transfer runtime has not started yet.");
        }

        return _app.Services.GetRequiredService<T>();
    }

    public async ValueTask DisposeAsync()
    {
        if (_app is not null)
        {
            await _app.StopAsync();
            await _app.DisposeAsync();
        }
    }
}
