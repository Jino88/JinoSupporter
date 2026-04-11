using Microsoft.AspNetCore.Components.Server.Circuits;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Scoped per Blazor Server circuit. Tracks circuit lifecycle and allows
/// the page to register a display name for the connected user.
/// </summary>
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
        _usersService.AddUser(circuit.Id, "Anonymous");
        return Task.CompletedTask;
    }

    public override Task OnCircuitClosedAsync(Circuit circuit, CancellationToken cancellationToken)
    {
        _usersService.RemoveUser(circuit.Id);
        return Task.CompletedTask;
    }

    /// <summary>Called from the layout to set the user's display name.</summary>
    public void SetName(string name)
    {
        if (!string.IsNullOrWhiteSpace(_circuitId) && !string.IsNullOrWhiteSpace(name))
            _usersService.UpdateName(_circuitId, name.Trim());
    }
}
