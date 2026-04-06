using System;
using System.Collections.Generic;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace QuickShareClone.Server;

public sealed class DiscoveryBroadcastService : BackgroundService
{
    private readonly DeviceIdentityService _deviceIdentityService;
    private readonly DiscoveryOptions _options;
    private readonly ILogger<DiscoveryBroadcastService> _logger;

    public DiscoveryBroadcastService(
        DeviceIdentityService deviceIdentityService,
        Microsoft.Extensions.Options.IOptions<DiscoveryOptions> options,
        ILogger<DiscoveryBroadcastService> logger)
    {
        _deviceIdentityService = deviceIdentityService;
        _options = options.Value;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        using var udpClient = new UdpClient();
        udpClient.EnableBroadcast = true;

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                var device = _deviceIdentityService.GetCurrentDevice() with { LastUpdatedAt = DateTimeOffset.UtcNow };
                var payload = JsonSerializer.SerializeToUtf8Bytes(device);
                foreach (var endpoint in GetBroadcastEndpoints())
                {
                    await udpClient.SendAsync(payload, payload.Length, endpoint);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to broadcast discovery beacon.");
            }

            await Task.Delay(TimeSpan.FromSeconds(_options.BroadcastIntervalSeconds), stoppingToken);
        }
    }

    private IReadOnlyCollection<IPEndPoint> GetBroadcastEndpoints()
    {
        var endpoints = new HashSet<IPEndPoint>(new IpEndPointComparer())
        {
            new(IPAddress.Broadcast, _options.BroadcastPort)
        };

        var addresses = NetworkInterface.GetAllNetworkInterfaces()
            .Where(x => x.OperationalStatus == OperationalStatus.Up)
            .SelectMany(x => x.GetIPProperties().UnicastAddresses)
            .Where(x => x.Address.AddressFamily == AddressFamily.InterNetwork && x.IPv4Mask is not null);

        foreach (var address in addresses)
        {
            endpoints.Add(new IPEndPoint(GetBroadcastAddress(address.Address, address.IPv4Mask!), _options.BroadcastPort));
        }

        return endpoints.ToArray();
    }

    private static IPAddress GetBroadcastAddress(IPAddress address, IPAddress subnetMask)
    {
        var addressBytes = address.GetAddressBytes();
        var maskBytes = subnetMask.GetAddressBytes();
        var broadcastBytes = new byte[addressBytes.Length];

        for (var i = 0; i < broadcastBytes.Length; i++)
        {
            broadcastBytes[i] = (byte)(addressBytes[i] | ~maskBytes[i]);
        }

        return new IPAddress(broadcastBytes);
    }

    private sealed class IpEndPointComparer : IEqualityComparer<IPEndPoint>
    {
        public bool Equals(IPEndPoint? x, IPEndPoint? y)
            => x?.Address.Equals(y?.Address) == true && x.Port == y?.Port;

        public int GetHashCode(IPEndPoint obj) => HashCode.Combine(obj.Address, obj.Port);
    }
}
