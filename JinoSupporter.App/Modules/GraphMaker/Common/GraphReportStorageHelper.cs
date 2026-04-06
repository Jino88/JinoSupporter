using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace GraphMaker;

public sealed class RawTableData
{
    public List<string> Columns { get; set; } = new();
    public List<List<string>> Rows { get; set; } = new();
}

public static class GraphReportStorageHelper
{
    public static RawTableData CaptureRawTableData(DataTable? table)
    {
        if (table == null)
        {
            return new RawTableData();
        }

        return new RawTableData
        {
            Columns = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToList(),
            Rows = table.Rows.Cast<DataRow>()
                .Select(row => table.Columns.Cast<DataColumn>()
                    .Select(column => row[column]?.ToString() ?? string.Empty)
                    .ToList())
                .ToList()
        };
    }

    public static DataTable BuildTableFromRawData(RawTableData? rawData)
    {
        var table = new DataTable();
        if (rawData == null)
        {
            return table;
        }

        foreach (string columnName in rawData.Columns)
        {
            table.Columns.Add(columnName);
        }

        foreach (List<string> rowValues in rawData.Rows)
        {
            DataRow row = table.NewRow();
            for (int i = 0; i < table.Columns.Count && i < rowValues.Count; i++)
            {
                row[i] = rowValues[i];
            }

            table.Rows.Add(row);
        }

        return table;
    }
}
