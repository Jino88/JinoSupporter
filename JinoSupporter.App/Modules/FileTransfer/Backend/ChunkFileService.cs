using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace QuickShareClone.Server;

public sealed class ChunkFileService
{
    private readonly UploadOptions _options;
    private readonly ILogger<ChunkFileService> _logger;
    private readonly string _rootPath;

    public ChunkFileService(IOptions<UploadOptions> options, ILogger<ChunkFileService> logger)
    {
        _options = options.Value;
        _logger = logger;
        _rootPath = Path.GetFullPath(_options.RootPath, AppContext.BaseDirectory);
        Directory.CreateDirectory(GetTempRoot());
        Directory.CreateDirectory(GetCompletedRoot());
    }

    public string GetChunkDirectory(string fileId)
    {
        var path = Path.Combine(GetTempRoot(), fileId);
        Directory.CreateDirectory(path);
        return path;
    }

    public string GetChunkPath(string fileId, int chunkIndex)
        => Path.Combine(GetChunkDirectory(fileId), $"chunk_{chunkIndex:D8}.part");

    public IReadOnlyCollection<int> GetReceivedChunks(string fileId)
    {
        var directory = GetChunkDirectory(fileId);
        return Directory.EnumerateFiles(directory, "chunk_*.part")
            .Select(Path.GetFileNameWithoutExtension)
            .Select(name => name?["chunk_".Length..])
            .Select(x => int.TryParse(x, out var value) ? (int?)value : null)
            .Where(x => x.HasValue)
            .Select(x => x!.Value)
            .OrderBy(x => x)
            .ToArray();
    }

    public async Task<long> SaveChunkAsync(string fileId, int chunkIndex, Stream input, CancellationToken cancellationToken)
    {
        var chunkPath = GetChunkPath(fileId, chunkIndex);
        if (File.Exists(chunkPath))
        {
            return new FileInfo(chunkPath).Length;
        }

        await using var output = new FileStream(chunkPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, true);
        await input.CopyToAsync(output, cancellationToken);
        await output.FlushAsync(cancellationToken);
        return output.Length;
    }

    public async Task<string> MergeChunksAsync(
        string fileId,
        string fileName,
        int totalChunks,
        string? destinationDirectory,
        CancellationToken cancellationToken)
    {
        var safeFileName = Path.GetFileName(fileName);
        var targetDirectory = string.IsNullOrWhiteSpace(destinationDirectory)
            ? GetCompletedRoot()
            : Path.GetFullPath(destinationDirectory);

        Directory.CreateDirectory(targetDirectory);
        var outputPath = EnsureUniqueFilePath(Path.Combine(targetDirectory, safeFileName));

        await using var output = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, true);

        for (var index = 0; index < totalChunks; index++)
        {
            var chunkPath = GetChunkPath(fileId, index);
            if (!File.Exists(chunkPath))
            {
                throw new InvalidOperationException($"Missing chunk {index} for fileId '{fileId}'.");
            }

            await using var input = new FileStream(chunkPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
            await input.CopyToAsync(output, cancellationToken);
        }

        TryDeleteDirectory(GetChunkDirectory(fileId));
        return outputPath;
    }

    public void CleanupExpiredChunks()
    {
        var expiration = DateTime.UtcNow.AddHours(-_options.ChunkExpirationHours);
        foreach (var directory in Directory.EnumerateDirectories(GetTempRoot()))
        {
            var info = new DirectoryInfo(directory);
            if (info.LastWriteTimeUtc < expiration)
            {
                TryDeleteDirectory(directory);
            }
        }
    }

    private string GetTempRoot() => Path.Combine(_rootPath, "temp");

    private string GetCompletedRoot() => Path.Combine(_rootPath, "completed");

    private void TryDeleteDirectory(string directory)
    {
        try
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to delete directory {Directory}", directory);
        }
    }

    private static string EnsureUniqueFilePath(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return filePath;
        }

        var directory = Path.GetDirectoryName(filePath)!;
        var name = Path.GetFileNameWithoutExtension(filePath);
        var extension = Path.GetExtension(filePath);
        var counter = 1;

        while (true)
        {
            var candidate = Path.Combine(directory, $"{name} ({counter}){extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            counter++;
        }
    }
}
