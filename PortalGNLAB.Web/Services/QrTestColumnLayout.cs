namespace PortalGNLAB.Web.Services;

/// <summary>
/// Left-to-right column order for one QR test table, shared by the data browser and the
/// workbook export so a column never sits in a different place in the file than on screen.
/// </summary>
public static class QrTestColumnLayout
{
    /// <summary>Marks the browser's derived test-date column inside a layout.</summary>
    public const int TestDateIndex = -1;

    private static readonly HashSet<string> NumericTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "bigint", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinyint",
    };

    /// <summary>
    /// Column indexes in display order, following <see cref="QrTestTableDef.ColumnOrder"/>.
    /// Columns the table definition does not mention keep their database order and follow the
    /// named ones, so a column added to the DB later still shows up. Exactly one entry is
    /// <see cref="TestDateIndex"/>.
    /// </summary>
    public static IReadOnlyList<int> Build(QrTestTableDef table, IReadOnlyList<QrTestDataColumn> columns)
    {
        var layout = new List<int>(columns.Count + 1);
        var placed = new HashSet<int>();
        bool testDatePlaced = false;

        foreach (string name in table.ColumnOrder)
        {
            if (string.Equals(name, QrTestTables.TestDateToken, StringComparison.Ordinal))
            {
                if (!testDatePlaced)
                {
                    layout.Add(TestDateIndex);
                    testDatePlaced = true;
                }
                continue;
            }

            int index = FindColumnIndex(columns, name);
            if (index >= 0 && placed.Add(index))
                layout.Add(index);
        }

        // Default layout — and the fallback for a table with no declared order — puts the
        // derived test date first, right after the sticky QR group column.
        if (!testDatePlaced)
            layout.Insert(0, TestDateIndex);

        for (int index = 0; index < columns.Count; index++)
        {
            if (placed.Add(index))
                layout.Add(index);
        }

        return layout;
    }

    /// <summary>Loose match: case and punctuation are ignored, so "Product_ID" finds ProductID.</summary>
    public static int FindColumnIndex(IReadOnlyList<QrTestDataColumn> columns, string expectedName)
    {
        string normalizedExpected = NormalizeColumnName(expectedName);
        for (int i = 0; i < columns.Count; i++)
        {
            if (NormalizeColumnName(columns[i].Name) == normalizedExpected)
                return i;
        }
        return -1;
    }

    public static bool IsNumeric(string dataType) => NumericTypes.Contains(dataType);

    private static string NormalizeColumnName(string value) =>
        new(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
}
