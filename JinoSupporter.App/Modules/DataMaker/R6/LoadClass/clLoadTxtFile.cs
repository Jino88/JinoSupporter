using System.Data;
using System.IO;

namespace DataMaker.R6.LoadClass
{
    public class clLoadTxtFile : ILoadData
    {
        public async Task<DataTable> Load(string fPath, string[] Header, char delimiter = '\t')
        {
            var table = new DataTable();

            foreach (var header in Header)
                table.Columns.Add(header, typeof(string));

            using (var reader = new StreamReader(fPath))
            {
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var values = line.Split(delimiter);
                    table.Rows.Add(values);
                }
            }

            return table;
        }
    }
}
