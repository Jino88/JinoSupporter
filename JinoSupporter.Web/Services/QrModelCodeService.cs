using System.Text.Json;

namespace JinoSupporter.Web.Services;

/// <summary>One row of the QR model table: the code carried in the QR, and the model it means.</summary>
public sealed class QrModelCode
{
    public string Code { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}

/// <summary>
/// The model table behind the QR pages. A QR carries its model in characters 11-12
/// (J69NAG1GD0<b>37</b>00009 → "37"), which is the only model marker the three stations agree on:
/// BKTD spells the model one way in <c>Model</c> (3200.MSM-X316.L1.M1.V01.260922, tail changes
/// with the test program) and ALTD/GSTD another in <c>ST_Information</c>
/// (3200.MSM-X316B.L1.M1.JIG1), so neither string can be filtered on across pages.
/// Backed by <c>AppStoragePaths.RootDirectory\qr-model-codes.json</c>, next to webapp-paths.json,
/// so it survives a publish and is shared by every page that reads QR data.
/// </summary>
public sealed class QrModelCodeService
{
    /// <summary>1-based position of the model code inside a QR, and how many characters it runs.</summary>
    public const int CodeStart = 11;
    public const int CodeLength = 2;

    /// <summary>Sentinel used by the UI for "a QR whose code is not in this table".</summary>
    public const string UnmappedCode = "#other";

    private const string FileName = "qr-model-codes.json";

    private static readonly JsonSerializerOptions WriteOptions = new() { WriteIndented = true };
    private static readonly JsonSerializerOptions ReadOptions = new() { PropertyNameCaseInsensitive = true };

    private readonly object _gate = new();
    private List<QrModelCode> _cached = [];
    private DateTime _lastWriteUtc = DateTime.MinValue;

    public QrModelCodeService()
    {
        SettingsPath = AppStoragePaths.Combine(FileName);
        lock (_gate)
        {
            if (File.Exists(SettingsPath))
            {
                _cached = LoadNoLock();
                _lastWriteUtc = LastWriteUtcNoLock();
                return;
            }
        }

        // Seeded from what the line actually ran over 2026-05 .. 2026-09; the table is editable
        // because BKTD also carries rework and test QR under other codes (00, 01, 11, 45 ...).
        Save([
            new QrModelCode { Code = "36", Name = "MSM-G778B" },
            new QrModelCode { Code = "37", Name = "MSM-X316B" },
        ]);
    }

    public string SettingsPath { get; }

    /// <summary>The table as saved, reloaded whenever the file changed underneath.</summary>
    public IReadOnlyList<QrModelCode> Codes
    {
        get
        {
            lock (_gate)
            {
                DateTime writeUtc = LastWriteUtcNoLock();
                if (writeUtc != _lastWriteUtc)
                {
                    _cached = LoadNoLock();
                    _lastWriteUtc = writeUtc;
                }

                return _cached.Select(Clone).ToList();
            }
        }
    }

    public void Save(IEnumerable<QrModelCode>? codes)
    {
        List<QrModelCode> normalized = Normalize(codes);
        lock (_gate)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? ".");
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(normalized, WriteOptions));
            _cached = normalized;
            _lastWriteUtc = LastWriteUtcNoLock();
        }
    }

    /// <summary>The model code inside a QR, or empty when the value is too short to carry one.</summary>
    public static string ExtractCode(string? productId)
    {
        string value = (productId ?? string.Empty).Trim();
        return value.Length < CodeStart - 1 + CodeLength
            ? string.Empty
            : value.Substring(CodeStart - 1, CodeLength).ToUpperInvariant();
    }

    /// <summary>Model name for a QR, or null when its code is not in the table.</summary>
    public string? ResolveName(string? productId)
    {
        string code = ExtractCode(productId);
        if (code.Length == 0)
            return null;

        foreach (QrModelCode entry in Codes)
        {
            if (string.Equals(entry.Code, code, StringComparison.OrdinalIgnoreCase))
                return entry.Name;
        }

        return null;
    }

    /// <summary>Drops blank codes and keeps the first entry of a duplicated code, so the table
    /// cannot end up with two rows a lookup would have to choose between.</summary>
    private static List<QrModelCode> Normalize(IEnumerable<QrModelCode>? codes)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<QrModelCode>();
        foreach (QrModelCode entry in codes ?? [])
        {
            string code = (entry.Code ?? string.Empty).Trim().ToUpperInvariant();
            if (code.Length == 0 || !seen.Add(code))
                continue;

            result.Add(new QrModelCode { Code = code, Name = (entry.Name ?? string.Empty).Trim() });
        }

        return result.OrderBy(entry => entry.Code, StringComparer.Ordinal).ToList();
    }

    private List<QrModelCode> LoadNoLock()
    {
        try
        {
            string json = File.ReadAllText(SettingsPath);
            return Normalize(JsonSerializer.Deserialize<List<QrModelCode>>(json, ReadOptions));
        }
        catch
        {
            // A corrupt table must not take the QR pages down with it; they fall back to "All models".
            return [];
        }
    }

    private DateTime LastWriteUtcNoLock() =>
        File.Exists(SettingsPath) ? File.GetLastWriteTimeUtc(SettingsPath) : DateTime.MinValue;

    private static QrModelCode Clone(QrModelCode source) =>
        new() { Code = source.Code, Name = source.Name };
}
