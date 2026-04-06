using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JinoSupporter.App.Infrastructure.Shell;

public interface IWorkbenchModuleBackend
{
    string Target { get; }

    IReadOnlyList<ShellActionDefinition> Actions { get; }

    event Action? SnapshotChanged;

    object? GetSnapshot();

    Task<object?> InvokeActionAsync(string action);
}
