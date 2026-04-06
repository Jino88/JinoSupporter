using System.IO;
using CustomKeyboardCSharp.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace CustomKeyboardCSharp.Services;

public sealed class GoogleDriveSyncService
{
    private static readonly string[] Scopes =
    [
        DriveService.Scope.DriveFile
    ];

    private readonly string _tokenDirectory;

    public GoogleDriveSyncService()
    {
        _tokenDirectory = CustomKeyboardPathResolver.GetAppDataPath("GoogleDriveTokens");
        Directory.CreateDirectory(_tokenDirectory);
    }

    public bool HasCachedLogin() => Directory.EnumerateFiles(_tokenDirectory, "*", SearchOption.AllDirectories).Any();

    public async Task<GoogleDriveSession> SignInAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        await AuthorizeAsync(settings, cancellationToken);
        return new GoogleDriveSession("Connected", true);
    }

    public async Task SignOutAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        if (HasCachedLogin())
        {
            try
            {
                var credential = await AuthorizeAsync(settings, cancellationToken);
                await credential.RevokeTokenAsync(cancellationToken);
            }
            catch
            {
                // Best effort: local token cleanup below still disconnects the desktop app.
            }
        }

        if (Directory.Exists(_tokenDirectory))
        {
            Directory.Delete(_tokenDirectory, recursive: true);
        }
        Directory.CreateDirectory(_tokenDirectory);
    }

    public async Task<UploadResult> UploadDatabaseAsync(
        AppSettings settings,
        string localDbPath,
        CancellationToken cancellationToken)
    {
        return await UploadNamedDatabaseAsync(settings, localDbPath, DesktopDbFileName, cancellationToken);
    }

    public async Task<UploadResult> UploadNamedDatabaseAsync(
        AppSettings settings,
        string localDbPath,
        string remoteFileName,
        CancellationToken cancellationToken)
    {
        var driveService = await CreateDriveServiceAsync(settings, cancellationToken);
        var folder = await GetOrCreateSyncFolderAsync(driveService, cancellationToken);
        var existing = await FindSyncFileAsync(driveService, remoteFileName, folder.Id!, cancellationToken);

        await using var contentStream = new FileStream(localDbPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        if (existing is null)
        {
            var metadata = new Google.Apis.Drive.v3.Data.File
            {
                Name = remoteFileName,
                Parents = [folder.Id]
            };
            var createRequest = driveService.Files.Create(metadata, contentStream, "application/octet-stream");
            createRequest.Fields = "id";
            await createRequest.UploadAsync(cancellationToken);
            return new UploadResult(false);
        }

        var updateMetadata = new Google.Apis.Drive.v3.Data.File
        {
            Name = remoteFileName
        };
        var updateRequest = driveService.Files.Update(updateMetadata, existing.Id, contentStream, "application/octet-stream");
        await updateRequest.UploadAsync(cancellationToken);
        return new UploadResult(true);
    }

    public async Task<bool> DownloadDatabaseAsync(AppSettings settings, string localDbPath, CancellationToken cancellationToken)
    {
        return await DownloadNamedDatabaseAsync(settings, localDbPath, DesktopDbFileName, cancellationToken);
    }

    public async Task<bool> DownloadNamedDatabaseAsync(
        AppSettings settings,
        string localDbPath,
        string remoteFileName,
        CancellationToken cancellationToken)
    {
        var driveService = await CreateDriveServiceAsync(settings, cancellationToken);
        var folder = await FindSyncFolderAsync(driveService, cancellationToken);
        if (folder is null)
        {
            return false;
        }
        var existing = await FindSyncFileAsync(driveService, remoteFileName, folder.Id!, cancellationToken);
        if (existing is null)
        {
            return false;
        }

        await using var memory = new FileStream(localDbPath, FileMode.Create, FileAccess.Write, FileShare.None);
        var getRequest = driveService.Files.Get(existing.Id);
        await getRequest.DownloadAsync(memory, cancellationToken);
        return true;
    }

    public async Task<GoogleDriveSession?> GetSessionAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        if (!HasCachedLogin())
        {
            return null;
        }

        try
        {
            var credential = await AuthorizeAsync(settings, cancellationToken);
            return new GoogleDriveSession("Connected", true);
        }
        catch
        {
            return null;
        }
    }

    private async Task<UserCredential> AuthorizeAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.GoogleClientId) || string.IsNullOrWhiteSpace(settings.GoogleClientSecret))
        {
            throw new InvalidOperationException("Add Google Desktop OAuth client ID and client secret in Settings first.");
        }

        var secrets = new ClientSecrets
        {
            ClientId = settings.GoogleClientId.Trim(),
            ClientSecret = settings.GoogleClientSecret.Trim()
        };

        return await GoogleWebAuthorizationBroker.AuthorizeAsync(
            secrets,
            Scopes,
            "desktop-user",
            cancellationToken,
            new FileDataStore(_tokenDirectory, true));
    }

    private async Task<DriveService> CreateDriveServiceAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        var credential = await AuthorizeAsync(settings, cancellationToken);
        return new DriveService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName
        });
    }

    private static async Task<Google.Apis.Drive.v3.Data.File?> FindSyncFolderAsync(DriveService driveService, CancellationToken cancellationToken)
    {
        var listRequest = driveService.Files.List();
        listRequest.Q = $"mimeType = 'application/vnd.google-apps.folder' and name = '{SyncFolderName}' and 'root' in parents and trashed = false";
        listRequest.Fields = "files(id,name,modifiedTime)";
        var result = await listRequest.ExecuteAsync(cancellationToken);
        return result.Files?.FirstOrDefault();
    }

    private static async Task<Google.Apis.Drive.v3.Data.File> GetOrCreateSyncFolderAsync(DriveService driveService, CancellationToken cancellationToken)
    {
        var existing = await FindSyncFolderAsync(driveService, cancellationToken);
        if (existing is not null)
        {
            return existing;
        }

        var metadata = new Google.Apis.Drive.v3.Data.File
        {
            Name = SyncFolderName,
            MimeType = "application/vnd.google-apps.folder",
            Parents = ["root"]
        };
        var createRequest = driveService.Files.Create(metadata);
        createRequest.Fields = "id,name,modifiedTime";
        return await createRequest.ExecuteAsync(cancellationToken);
    }

    private static async Task<Google.Apis.Drive.v3.Data.File?> FindSyncFileAsync(
        DriveService driveService,
        string fileName,
        string folderId,
        CancellationToken cancellationToken)
    {
        var listRequest = driveService.Files.List();
        listRequest.Q = $"name = '{fileName}' and '{folderId}' in parents and trashed = false";
        listRequest.Fields = "files(id,name,modifiedTime)";
        var result = await listRequest.ExecuteAsync(cancellationToken);
        return result.Files?
            .OrderByDescending(file => file.ModifiedTimeDateTimeOffset ?? DateTimeOffset.MinValue)
            .FirstOrDefault();
    }

    public sealed record GoogleDriveSession(string Email, bool IsConnected);
    public sealed record UploadResult(bool Updated);

    public const string SyncFolderName = "CustomKeyboard Sync";
    public const string DesktopDbFileName = "customkeyboard_desktop.db";
    public const string AndroidDbFileName = "customkeyboard_android.db";
    private const string ApplicationName = "CustomKeyboardCSharp";
}
