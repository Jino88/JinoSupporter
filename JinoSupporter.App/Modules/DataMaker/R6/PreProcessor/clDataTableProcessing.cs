using System.Data;

namespace DataMaker.R6.PreProcessor
{

    public static class clDataTableProcessing
    {

        public static DataTable MergeDataTable(List<DataTable> sourceTables)
        {
            DataTable Result = sourceTables[0].Clone(); // Clone the structure of the first table

            foreach (var table in sourceTables)
            {
                Result.Merge(table);
                DataRow d = Result.NewRow();

                foreach (DataColumn col in Result.Columns)
                {
                    d[col.ColumnName] = DBNull.Value;
                }

                Result.Rows.Add(d);
            }

            return Result;

        }

        public static DataTable SumColumnsByGroup(DataTable table, string RefColumnsName, int startCol)
        {
            DataTable result = table.Clone();

            var groups = table.AsEnumerable()
                .GroupBy(r => r.Field<string>(RefColumnsName));

            foreach (var g in groups)
            {
                DataRow newRow = result.NewRow();

                // 그룹키
                newRow[RefColumnsName] = g.Key;

                // 문자열 컬럼은 첫 번째 값 그대로 복사
                foreach (DataColumn col in table.Columns)
                {
                    if (col.Ordinal < startCol && col.ColumnName != RefColumnsName)
                    {
                        newRow[col.ColumnName] = g.First()[col];
                    }
                }

                // 숫자 컬럼은 Sum
                for (int col = startCol; col < table.Columns.Count; col++)
                {
                    double sum = g.Sum(r => r.IsNull(col) ? 0 : Convert.ToDouble(r[col]));
                    newRow[col] = sum;
                }

                result.Rows.Add(newRow);
            }

            return result;
        }

        public static void RemoveColumns(DataTable table, List<string> columnsToRemove)
        {
            // 뒤에서부터 제거해야 인덱스 꼬이지 않음
            foreach (string colName in columnsToRemove)
            {
                if (table.Columns.Contains(colName))
                {
                    table.Columns.Remove(colName);
                }
            }
        }
    }
}
