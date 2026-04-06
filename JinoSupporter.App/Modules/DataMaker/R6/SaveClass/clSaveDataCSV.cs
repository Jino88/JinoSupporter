using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Text;

namespace DataMaker.R6.SaveClass
{
    public class clSaveData
    {
        protected IEnumerable<object[]> ToRows(DataTable dt)
        {
            // 헤더
            yield return dt.Columns.Cast<DataColumn>().Select(c => (object)c.ColumnName).ToArray();

            // 데이터
            foreach (DataRow row in dt.Rows)
                yield return row.ItemArray;
        }
    }
    public class clSaveDataCSV : clSaveData, ISaveData
    {

        public Task Save(DataTable Table, string path)
        {
            if (Table == null) { return null; }
            else
            {
                if (string.IsNullOrWhiteSpace(path))
                    throw new ArgumentException("파일 경로를 지정해야 합니다.", nameof(path));

                // 파일이 열려있거나 쓰기 불가능한 경우 -1, -2 등을 붙여서 시도
                string finalPath = GetAvailableFilePath(path);

                using (var writer = new StreamWriter(finalPath, false, Encoding.UTF8))
                {
                    // 1. 헤더 작성
                    var columnNames = Table.Columns.Cast<DataColumn>()
                                                       .Select(c => c.ColumnName);
                    writer.WriteLine(string.Join(",", columnNames));

                    // 2. 데이터 행 작성
                    foreach (DataRow row in Table.Rows)
                    {
                        var fields = row.ItemArray.Select(field =>
                        {
                            if (field == null) return "";
                            string s = field.ToString();

                            // 쉼표나 따옴표 포함 시 CSV 규격에 맞게 처리
                            if (s.Contains(",") || s.Contains("\""))
                                s = "\"" + s.Replace("\"", "\"\"") + "\"";

                            return s;
                        });

                        writer.WriteLine(string.Join(",", fields));
                    }
                }
                return Task.CompletedTask;
            }


        }

        /// <summary>
        /// 파일이 이미 열려있거나 쓰기 불가능한 경우 -1, -2 등을 붙여서 사용 가능한 파일 경로 반환
        /// </summary>
        private string GetAvailableFilePath(string originalPath)
        {
            // 원본 파일 먼저 시도
            if (IsFileWritable(originalPath))
                return originalPath;

            // 파일명과 확장자 분리
            string directory = Path.GetDirectoryName(originalPath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);

            // -1, -2, -3... 형태로 시도
            int counter = 1;
            while (counter <= 100) // 최대 100개까지 시도
            {
                string newFileName = $"{fileNameWithoutExt}-{counter}{extension}";
                string newPath = string.IsNullOrEmpty(directory)
                    ? newFileName
                    : Path.Combine(directory, newFileName);

                if (IsFileWritable(newPath))
                    return newPath;

                counter++;
            }

            // 100개까지 시도했는데도 안되면 원본 경로 반환 (예외 발생하게 함)
            return originalPath;
        }

        /// <summary>
        /// 파일이 쓰기 가능한지 확인
        /// </summary>
        private bool IsFileWritable(string path)
        {
            // 파일이 존재하지 않으면 쓰기 가능
            if (!File.Exists(path))
                return true;

            // 파일이 존재하면 열려있는지 확인
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.None))
                {
                    // 쓰기 가능
                }
                return true;
            }
            catch (IOException)
            {
                // 파일이 열려있어서 쓰기 불가능
                return false;
            }
            catch
            {
                // 기타 오류 (권한 등)
                return false;
            }
        }

        public Task Save(List<DataTable> Tables, string path)
        {
            throw new NotImplementedException();
        }
    }

    public class clSaveDataExcel : clSaveData, ISaveData
    {
        public Task Save(DataTable Table, string path)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            if (Table == null) { return null; }
            else
            {
                if (string.IsNullOrWhiteSpace(path))
                    throw new ArgumentException("파일 경로를 지정해야 합니다.", nameof(path));

                FileInfo fileInfo = new FileInfo(path);
                if (fileInfo.Exists)
                    fileInfo.Delete();

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    string sheetName = string.IsNullOrEmpty(Table.TableName)
                        ? "Sheet" + Random.Shared.Next(1, 1000)
                        : Table.TableName;

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // 빠른 로드
                    worksheet.Cells["A1"].LoadFromArrays(ToRows(Table));

                    // 헤더만 AutoFit (속도 최적화)
                    worksheet.Cells[1, 1, 1, Table.Columns.Count].AutoFitColumns();

                    package.Save();
                }
                return Task.CompletedTask;
            }
        }
        public Task Save(List<DataTable> Tables, string path)
        {
            if (Tables == null || Tables.Count == 0)
                throw new ArgumentException("DataTable 목록이 비어 있거나 null입니다.", nameof(Tables));
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("파일 경로를 지정해야 합니다.", nameof(path));

            FileInfo fileInfo = new FileInfo(path);
            if (fileInfo.Exists)
                fileInfo.Delete();

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                HashSet<string> existingSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var table in Tables)
                {
                    string sheetName = string.IsNullOrEmpty(table.TableName)
                        ? "Sheet" + Random.Shared.Next(1, 1000)
                        : table.TableName;

                    // 중복 처리
                    string originalName = sheetName;
                    int counter = 1;
                    while (existingSheetNames.Contains(sheetName))
                    {
                        sheetName = $"{originalName}_{counter++}";
                    }
                    existingSheetNames.Add(sheetName);

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // 빠른 로드
                    worksheet.Cells["A1"].LoadFromArrays(ToRows(table));

                    // 헤더만 AutoFit
                    worksheet.Cells[1, 1, 1, table.Columns.Count].AutoFitColumns();
                }

                package.Save();
            }

            return Task.CompletedTask;
        }
    }
}
