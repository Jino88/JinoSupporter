using System.Collections.Concurrent;

namespace JinoSupporter.Web.Services;

public sealed record UserInfo(string CircuitId, string Name, DateTime ConnectedAt);

public sealed class ConnectedUsersService
{
    private readonly ConcurrentDictionary<string, UserInfo> _users = new();

    public event Action? Changed;

    public IReadOnlyList<UserInfo> Users =>
        [.. _users.Values.OrderBy(u => u.ConnectedAt)];

    public int Count => _users.Count;

    public void AddUser(string circuitId, string name = "Anonymous")
    {
        _users[circuitId] = new UserInfo(circuitId, name, DateTime.Now);
        Changed?.Invoke();
    }

    public void UpdateName(string circuitId, string name)
    {
        if (_users.TryGetValue(circuitId, out UserInfo? existing))
            _users[circuitId] = existing with { Name = name };
        Changed?.Invoke();
    }

    public void RemoveUser(string circuitId)
    {
        _users.TryRemove(circuitId, out _);
        Changed?.Invoke();
    }
}
