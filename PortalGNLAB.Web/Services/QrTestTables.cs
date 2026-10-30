namespace PortalGNLAB.Web.Services;

/// <summary>
/// One QR inspection table in TCDB. The three tables share a connection and the
/// ProductID key, but each spells its timestamp column and its pass/fail verdict
/// differently, so those differences are declared here instead of being guessed
/// inside the query builder.
/// </summary>
/// <param name="Key">Stable key used in routes and CSV names; also the bare table name.</param>
/// <param name="Station">Short station name shown in the summary grid (BAKO / AIRLEAK / GAUSS).</param>
/// <param name="MenuLabel">Nav menu and page title text.</param>
/// <param name="Table">Fully quoted table reference spliced into SQL. Never user input.</param>
/// <param name="DisplayTable">Human readable table name shown in the UI.</param>
/// <param name="TimeColumns">Timestamp column candidates, best first.</param>
/// <param name="PassValues">TestResult values that mean the unit passed.</param>
/// <param name="FailValues">TestResult values that mean the unit failed.</param>
public sealed record QrTestTableDef(
    string Key,
    string Station,
    string MenuLabel,
    string Table,
    string DisplayTable,
    IReadOnlyList<string> TimeColumns,
    IReadOnlyList<string> PassValues,
    IReadOnlyList<string> FailValues)
{
    /// <summary>A well formed QR carries exactly this many characters; anything else is scanner noise.</summary>
    public int ProductIdLength { get; init; } = 17;

    /// <summary>
    /// Left-to-right column order for the data browser. Names are matched loosely
    /// (case and punctuation are ignored); <see cref="QrTestTables.TestDateToken"/>
    /// marks where the browser's own derived test-date column goes. Any table column
    /// not listed here keeps its database order and is appended after these.
    /// Empty means "database order, test date first".
    /// </summary>
    public IReadOnlyList<string> ColumnOrder { get; init; } = [];

    public bool IsPass(string value) =>
        PassValues.Any(v => string.Equals(v, value.Trim(), StringComparison.OrdinalIgnoreCase));

    public bool IsFail(string value) =>
        FailValues.Any(v => string.Equals(v, value.Trim(), StringComparison.OrdinalIgnoreCase));
}

/// <summary>Registry of the three QR inspection tables in TCDB.</summary>
public static class QrTestTables
{
    public const string BakoKey = "BKTD";
    public const string AirleakKey = "ALTD";
    public const string GaussKey = "GSTD";

    /// <summary>Placeholder inside <see cref="QrTestTableDef.ColumnOrder"/> for the derived test-date column.</summary>
    public const string TestDateToken = "#TestDate";

    public static readonly QrTestTableDef Bako = new(
        BakoKey,
        "BAKO",
        "QR BAKO DATA",
        "[dbo].[BKTD]",
        "dbo.BKTD",
        ["TestTime"],
        ["PASS", "OK"],
        ["FAIL", "NG", "N/G"]);

    public static readonly QrTestTableDef Airleak = new(
        AirleakKey,
        "AIRLEAK",
        "QR AIRLEAK DATA",
        "[dbo].[ALTD]",
        "dbo.ALTD",
        // The tester PC's clock has been wrong (G778 units stamped a month early), so the
        // date a row reached the database is the one dates and retests are judged by.
        ["ImportedAt", "TestTime"],
        ["OK", "PASS"],
        // "AL" is what the air leak tester writes for a leaking unit.
        ["AL", "NG", "FAIL", "N/G"])
    {
        ColumnOrder =
        [
            "id",
            TestDateToken,
            "ImportedAt",
            "TestTime",
            "ST_Information",
            "ProductID",
            "TestResult",
            "description",
            // Every pressure reading, leak first.
            "leak_pressure_val",
            "leak_pressure_unit",
            "input_pressure_val",
            "input_pressure_unit",
        ],
    };

    public static readonly QrTestTableDef Gauss = new(
        GaussKey,
        "GAUSS",
        "QR GAUSS DATA",
        "[dbo].[GSTD]",
        "dbo.GSTD",
        // ImportedAt first, as for ALTD; GSTD does not have that column yet, so today this
        // falls through to OTIME (it has no TestTime either) and switches by itself once added.
        ["ImportedAt", "OTIME", "TestTime"],
        // The gauss tester stores its verdict as a flag: 1 = pass, 0 = fail.
        ["1", "OK", "PASS"],
        ["0", "NG", "FAIL", "N/G"])
    {
        ColumnOrder =
        [
            TestDateToken,
            "ST_Information",
            "ProductID",
            "OTIME",
            "TestResult",
            // Each reading sits next to its own spec window, low before high.
            "Value",
            "SpecLow",
            "SpecHigh",
            "Value2",
            "SpecLow2",
            "SpecHigh2",
        ],
    };

    public static readonly IReadOnlyList<QrTestTableDef> All = [Bako, Airleak, Gauss];

    /// <summary>
    /// Maps a key to its definition. Unknown keys fall back to BAKO so a bad route can
    /// never reach the query builder with an attacker supplied table name.
    /// </summary>
    public static QrTestTableDef Resolve(string? key) =>
        All.FirstOrDefault(t => string.Equals(t.Key, key?.Trim(), StringComparison.OrdinalIgnoreCase)) ?? Bako;
}
