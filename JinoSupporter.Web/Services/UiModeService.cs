namespace JinoSupporter.Web.Services;

/// <summary>
/// Which shell the app renders in: the classic <c>MainLayout</c> or the redesigned
/// <c>InstrumentLayout</c>. Scoped, so the choice belongs to one browser circuit —
/// turning the new UI on never changes what anyone else sees.
///
/// <see cref="Routes"/> picks the layout from this and re-renders on <see cref="Changed"/>;
/// the value is mirrored into localStorage so it survives a full page reload.
/// </summary>
public sealed class UiModeService
{
    public const string StorageKey = "jino-ui-mode";

    private bool _newUi;

    public bool NewUi
    {
        get => _newUi;
        set
        {
            if (_newUi == value) return;
            _newUi = value;
            Changed?.Invoke();
        }
    }

    public event Action? Changed;
}
