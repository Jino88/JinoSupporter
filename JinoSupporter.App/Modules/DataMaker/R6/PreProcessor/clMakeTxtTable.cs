using DataMaker.R6.SQLProcess;
using System.Data;
using System.IO;

namespace DataMaker.R6.PreProcessor
{
    public class clMakeTxtTable
    {
        clSQLFileIO sql;
        public clMakeTxtTable(string dbPath)
        {
            sql = new clSQLFileIO(dbPath);
        }

        public void Make(string TableName, string txtPath, List<string> Columns)
        {
            DataTable txtTable = MakeDataTable(txtPath, Columns);
            sql.Writer.Write(TableName, txtTable);
        }

        public DataTable MakeDataTable(string txtPath, List<string> Columns)
        {
            DataTable table = new DataTable();

            table.BeginLoadData();

            using (var reader = new StreamReader(txtPath))
            {
                bool isFirstLine = true;
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // 탭으로 분리
                    string[] values = line.Split('\t');

                    // 첫 줄: 헤더 스킵 (컬럼은 Columns 파라미터로 이미 지정됨)
                    if (isFirstLine)
                    {
                        foreach (string col in Columns)
                            table.Columns.Add(col.Trim(), typeof(string));

                        isFirstLine = false;
                        continue; // 첫 줄 헤더 스킵
                    }

                    // 데이터 행 추가
                    DataRow row = table.NewRow();
                    for (int i = 0; i < values.Length && i < Columns.Count; i++)
                    {
                        // Raw data 로드 (normalization은 caller가 필요시 적용)
                        row[i] = values[i].Trim();
                    }

                    table.Rows.Add(row);
                }
            }

            table.EndLoadData();
            return table;
        }
    }
}
