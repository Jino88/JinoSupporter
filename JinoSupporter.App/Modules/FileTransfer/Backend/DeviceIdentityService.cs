using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using Microsoft.Extensions.Options;

namespace QuickShareClone.Server;

public sealed class DeviceIdentityService
{
    private readonly DiscoveryOptions _options;
    private readonly Lazy<DiscoveryDeviceInfo> _deviceInfo;

    public DeviceIdentityService(IOptions<DiscoveryOptions> options)
    {
        _options = options.Value;
        _deviceInfo = new Lazy<DiscoveryDeviceInfo>(CreateDeviceInfo);
    }

    public DiscoveryDeviceInfo GetCurrentDevice() => _deviceInfo.Value;

    private DiscoveryDeviceInfo CreateDeviceInfo()
    {
        var deviceId = Environment.MachineName.ToLowerInvariant();
        var urls = GetServerUrls();
        return new DiscoveryDeviceInfo(
            DeviceId: deviceId,
            DeviceName: Environment.MachineName,
            Platform: "windows",
            ServerUrls: urls,
            LastUpdatedAt: DateTimeOffset.UtcNow);
    }

    private static IReadOnlyCollection<string> GetServerUrls()
    {
        var addresses = NetworkInterface.GetAllNetworkInterfaces()
            .Where(x => x.OperationalStatus == OperationalStatus.Up)
            .SelectMany(x => x.GetIPProperties().UnicastAddresses)
            .Select(x => x.Address)
            .Where(x => x.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(x))
            .Select(x => $"http://{x}:5070")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (addresses.Length == 0)
        {
            return ["http://127.0.0.1:5070"];
        }

        return addresses;
    }
}
