using System.Collections.Concurrent;

namespace JinoSupporter.Web.Services;

public sealed record UserInfo(string CircuitId, string Username, string Name, DateTime ConnectedAt);

public sealed class ConnectedUsersService
{
    private readonly ConcurrentDictionary<string, UserInfo> _users = new();

    public event Action? Changed;

    public IReadOnlyList<UserInfo> Users =>
        [.. _users.Values.OrderBy(u => u.ConnectedAt)];

    public int Count => _users.Count;

    public void AddUser(string circuitId, string username = "", string name = "Anonymous")
    {
        _users[circuitId] = new UserInfo(circuitId, username, name, DateTime.Now);
        Changed?.Invoke();
    }

    public void UpdateName(string circuitId, string name)
    {
        if (_users.TryGetValue(circuitId, out UserInfo? existing))
            _users[circuitId] = existing with { Name = name };
        Changed?.Invoke();
    }

    public void UpdateNameByUsername(string username, string name)
    {
        if (string.IsNullOrEmpty(username)) return;
        bool changed = false;
        foreach (var kv in _users)
        {
            if (kv.Value.Username.Equals(username, StringComparison.OrdinalIgnoreCase))
            {
                _users[kv.Key] = kv.Value with { Name = name };
                changed = true;
            }
        }
        if (changed) Changed?.Invoke();
    }

    public void RemoveUser(string circuitId)
    {
        _users.TryRemove(circuitId, out _);
        Changed?.Invoke();
    }
}
