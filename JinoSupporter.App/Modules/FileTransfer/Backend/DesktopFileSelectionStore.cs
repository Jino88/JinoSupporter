namespace QuickShareClone.Server;

public sealed class DesktopFileSelectionStore
{
    private readonly object _gate = new();
    private readonly Dictionary<string, DesktopNativeSelection> _selections = new(StringComparer.Ordinal);

    public DesktopNativeSelection Save(IReadOnlyCollection<DesktopSelectedFile> files)
    {
        var selection = new DesktopNativeSelection(
            SelectionId: Guid.NewGuid().ToString("N"),
            Files: files,
            CreatedAt: DateTimeOffset.UtcNow);

        lock (_gate)
        {
            _selections[selection.SelectionId] = selection;
        }

        return selection;
    }

    public DesktopNativeSelection? Find(string selectionId)
    {
        lock (_gate)
        {
            return _selections.GetValueOrDefault(selectionId);
        }
    }

    public void Remove(string selectionId)
    {
        lock (_gate)
        {
            _selections.Remove(selectionId);
        }
    }
}
