using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
namespace QuickShareClone.Server;

public sealed class CleanupService : BackgroundService
{
    private readonly ChunkFileService _chunkFileService;
    private readonly ILogger<CleanupService> _logger;

    public CleanupService(ChunkFileService chunkFileService, ILogger<CleanupService> logger)
    {
        _chunkFileService = chunkFileService;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                _chunkFileService.CleanupExpiredChunks();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Background cleanup failed.");
            }

            await Task.Delay(TimeSpan.FromMinutes(30), stoppingToken);
        }
    }
}
