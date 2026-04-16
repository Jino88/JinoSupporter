using Microsoft.AspNetCore.Components.Server.Circuits;

namespace JinoSupporter.Web.Services;

public sealed class UserCircuitHandler : CircuitHandler
{
    private readonly ConnectedUsersService _usersService;
    private string? _circuitId;

    public UserCircuitHandler(ConnectedUsersService usersService)
    {
        _usersService = usersService;
    }

    public override Task OnCircuitOpenedAsync(Circuit circuit, CancellationToken cancellationToken)
    {
        _circuitId = circuit.Id;
        _usersService.AddUser(circuit.Id);
        return Task.CompletedTask;
    }

    public override Task OnCircuitClosedAsync(Circuit circuit, CancellationToken cancellationToken)
    {
        _usersService.RemoveUser(circuit.Id);
        return Task.CompletedTask;
    }

    /// <summary>Called from the layout after auth state is resolved.</summary>
    public void Register(string username, string displayName)
    {
        if (string.IsNullOrWhiteSpace(_circuitId)) return;
        string name = string.IsNullOrWhiteSpace(displayName) ? username : displayName;
        _usersService.AddUser(_circuitId, username, name);
    }
}
