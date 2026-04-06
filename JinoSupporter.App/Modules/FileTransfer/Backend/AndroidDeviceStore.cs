using System.Collections.Concurrent;

namespace QuickShareClone.Server;

public sealed class AndroidDeviceStore
{
    private readonly ConcurrentDictionary<string, AndroidConnectedDevice> _devices = new();

    public void Register(AndroidDeviceRegistrationRequest request)
    {
        var normalizedUrl = request.ReceiveUrl.Trim().TrimEnd('/');
        var device = new AndroidConnectedDevice(
            request.DeviceId,
            request.DeviceName,
            request.Platform,
            normalizedUrl,
            DateTimeOffset.UtcNow);

        _devices.AddOrUpdate(request.DeviceId, device, (_, _) => device);
    }

    public AndroidConnectedDevice? Find(string deviceId) =>
        _devices.TryGetValue(deviceId, out var device) ? device : null;

    public IReadOnlyCollection<AndroidConnectedDevice> List() =>
        _devices.Values
            .OrderByDescending(x => x.LastSeenAt)
            .ToArray();
}
