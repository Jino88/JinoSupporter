using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JinoSupporter.App.Infrastructure.Shell;

public sealed class DelegateModuleBackend<TModule> : IWorkbenchModuleBackend
    where TModule : class
{
    private readonly Func<TModule> _getModule;
    private readonly Func<TModule, object?> _getSnapshot;
    private readonly Func<TModule, string, Task<object?>> _invokeActionAsync;
    private readonly Action<TModule, Action>? _subscribe;
    private readonly Action<TModule, Action>? _unsubscribe;
    private TModule? _subscribedModule;

    public DelegateModuleBackend(
        string target,
        IReadOnlyList<ShellActionDefinition> actions,
        Func<TModule> getModule,
        Func<TModule, object?> getSnapshot,
        Func<TModule, string, Task<object?>> invokeActionAsync,
        Action<TModule, Action>? subscribe = null,
        Action<TModule, Action>? unsubscribe = null)
    {
        Target = target;
        Actions = actions;
        _getModule = getModule;
        _getSnapshot = getSnapshot;
        _invokeActionAsync = invokeActionAsync;
        _subscribe = subscribe;
        _unsubscribe = unsubscribe;
    }

    public string Target { get; }

    public IReadOnlyList<ShellActionDefinition> Actions { get; }

    public event Action? SnapshotChanged;

    public object? GetSnapshot()
    {
        TModule module = GetModule();
        return _getSnapshot(module);
    }

    public Task<object?> InvokeActionAsync(string action)
    {
        TModule module = GetModule();
        return _invokeActionAsync(module, action);
    }

    private TModule GetModule()
    {
        TModule module = _getModule();
        if (!ReferenceEquals(_subscribedModule, module))
        {
            if (_subscribedModule is not null && _unsubscribe is not null)
            {
                _unsubscribe(_subscribedModule, HandleSnapshotChanged);
            }

            _subscribedModule = module;

            if (_subscribe is not null)
            {
                _subscribe(_subscribedModule, HandleSnapshotChanged);
            }
        }

        return module;
    }

    private void HandleSnapshotChanged()
    {
        SnapshotChanged?.Invoke();
    }
}
