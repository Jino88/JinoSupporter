using System.Text.Json;

namespace PortalGNLAB.Web.Services;

/// <summary>What the QR pages remember between visits.</summary>
public sealed class QrViewPreferences
{
    /// <summary>Max rows last picked on a QR data page; 0 means "never picked" and
    /// <see cref="QrViewPreferenceService.AutoValue"/> means "fit the selection".</summary>
    public int MaxRows { get; set; }

    /// <summary>Max QR last picked on QR SUMMARY or the GAUSS scatter; 0 means "never picked".</summary>
    public int MaxQr { get; set; }
}

/// <summary>
/// Remembers the row caps the QR pages were last set to, so a pick made on one page holds for the
/// others and survives a restart — the four QR menus otherwise each reset to the configured
/// default on every visit, and the cap has to be re-picked before anything can be read.
/// Backed by <c>AppStoragePaths.RootDirectory\qr-view-prefs.json</c>, next to webapp-paths.json.
/// </summary>
/// <remarks>
/// Rows and QR are kept apart on purpose: a data page counts test rows while SUMMARY and the
/// scatter count QR, and one number would mean two different amounts of data.
/// </remarks>
public sealed class QrViewPreferenceService
{
    private const string FileName = "qr-view-prefs.json";

    /// <summary>Stored instead of a number when the page should fit the cap to the selection.</summary>
    public const int AutoValue = -1;

    private static readonly JsonSerializerOptions WriteOptions = new() { WriteIndented = true };
    private static readonly JsonSerializerOptions ReadOptions = new() { PropertyNameCaseInsensitive = true };

    private readonly object _gate = new();
    private QrViewPreferences _cached;

    public QrViewPreferenceService()
    {
        SettingsPath = AppStoragePaths.Combine(FileName);
        _cached = Load();
    }

    public string SettingsPath { get; }

    /// <summary>Last picked row cap, or 0 when the user has never picked one.</summary>
    public int MaxRows
    {
        get { lock (_gate) return _cached.MaxRows; }
    }

    /// <summary>Last picked QR cap, or 0 when the user has never picked one.</summary>
    public int MaxQr
    {
        get { lock (_gate) return _cached.MaxQr; }
    }

    public void SetMaxRows(int value) => Update(prefs => prefs.MaxRows = Normalize(value));

    public void SetMaxQr(int value) => Update(prefs => prefs.MaxQr = Normalize(value));

    /// <summary>The remembered value, or <paramref name="fallback"/> when nothing was ever picked.</summary>
    public static int Resolve(int remembered, int fallback) => remembered > 0 ? remembered : fallback;

    private void Update(Action<QrViewPreferences> change)
    {
        lock (_gate)
        {
            var next = new QrViewPreferences { MaxRows = _cached.MaxRows, MaxQr = _cached.MaxQr };
            change(next);
            if (next.MaxRows == _cached.MaxRows && next.MaxQr == _cached.MaxQr)
                return;

            _cached = next;
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? ".");
                File.WriteAllText(SettingsPath, JsonSerializer.Serialize(next, WriteOptions));
            }
            catch
            {
                // A preference that cannot be written is not worth failing a page load over; the
                // value still holds for this run of the app.
            }
        }
    }

    private QrViewPreferences Load()
    {
        try
        {
            if (!File.Exists(SettingsPath))
                return new QrViewPreferences();

            var loaded = JsonSerializer.Deserialize<QrViewPreferences>(File.ReadAllText(SettingsPath), ReadOptions)
                ?? new QrViewPreferences();
            return new QrViewPreferences
            {
                MaxRows = Normalize(loaded.MaxRows),
                MaxQr = Normalize(loaded.MaxQr),
            };
        }
        catch
        {
            return new QrViewPreferences();
        }
    }

    // 0 keeps "never picked" meaningful and -1 is the auto sentinel; a real cap is clamped to
    // what the pages are allowed to ask for.
    private static int Normalize(int value) => value switch
    {
        AutoValue => AutoValue,
        <= 0 => 0,
        _ => Math.Min(value, 20000),
    };
}
