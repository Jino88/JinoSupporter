using DataMaker.Logger;
using DataMaker.R6.LoadClass;
using DataMaker.R6.SQLProcess;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DataMaker.R6.FetchDataBMES
{
    public class BMESOption
    {
        public DateTime StartDate, EndDate;
        public InfoID InfoId;

        public void SetLoginID(string ID, string Password)
        {
            InfoId = new InfoID(ID, Password);
        }


    }

    public class clFetchBMESNGDATA
    {
        private const string OriginalTableMetaTableName = "__DataMakerMeta";
        private const string OriginalTableUpdatedAtKey = "OriginalTableUpdatedAt";
        private DataTable DataTableBMES;

        public clFetchBMESNGDATA()
        {
        }

        public async Task GetDataFromWebAsync(BMESOption Option)
        {
            if (Option == null)
            {
                throw new ArgumentNullException(nameof(Option), "BMESOption cannot be null.");
            }

            var stopwatch = Stopwatch.StartNew();
            DataTableBMES = await GetDataFromWeb(Option);
            stopwatch.Stop();

            int rowCount = DataTableBMES?.Rows.Count ?? 0;
            clLogger.LogImportant($"GetDataFromWebAsync loaded {rowCount:N0} rows in {stopwatch.Elapsed.TotalSeconds:F2}s");

        }
        private async Task<DataTable> GetDataFromWeb(BMESOption Option)
        {
            DateTime StartDate = Option.StartDate;
            DateTime EndDate = Option.EndDate;

            if (Option.InfoId == null || string.IsNullOrEmpty(Option.InfoId.LoginID) || string.IsNullOrEmpty(Option.InfoId.Password))
            {
                MessageBox.Show("자격 증명이 설정되지 않았습니다. 먼저 BMES 설정을 완료해주세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            try
            {
                clFetchBMES dataFetcher = new clFetchBMES(
                    Option.InfoId.LoginID,
                    Option.InfoId.Password,
                    StartDate.ToString("yyyy-MM-dd"),
                    EndDate.ToString("yyyy-MM-dd"));

                System.Data.DataTable d = await dataFetcher.RunAsync();
                
                return d;

            }

            catch (Exception ex)
            {
                MessageBox.Show($"Failed Load Data From BMES {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

        }
        public async Task<string> SaveBMESDataToDB(bool isAdd)
        {
            var totalStopwatch = Stopwatch.StartNew();
            string dbFilePath = null;

            if (isAdd)
            {
                // 기존 DB에 추가: OpenFileDialog 사용
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "db Files|*.db",
                    Title = "Select existing DB to add data"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    dbFilePath = openFileDialog.FileName;
                }
            }
            else
            {
                // 새 DB 생성: SaveFileDialog 사용
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "db Files|*.db",
                    Title = "Save BMES Data",
                    FileName = "" + DateTime.Now.ToString("yyMMdd_HHmmss") + ".db"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    dbFilePath = saveFileDialog.FileName;
                }
            }

            if (string.IsNullOrEmpty(dbFilePath))
            {
                clLogger.Log("BMES Data Save Canceled.");
                return null;
            }
            else
            {
                clSQLFileIO sql = new clSQLFileIO(dbFilePath);

/*                clSQLiteLoader l = new clSQLiteLoader(saveFileDialog.FileName);
                clSQLiteDataProcessing dp = new clSQLiteDataProcessing(l, new clSQLiteReader(l), null);
                clSQLiteWriter c = new clSQLiteWriter(l, dp);*/

                if (!isAdd)
                {
                    // 새 DB 생성: 새 데이터도 중복 제거 필요 (BMES에서 받은 데이터 자체에 중복 있을 수 있음)
                    clLogger.Log("Creating new DB. Removing duplicates from new data...");
                    var duplicateStopwatch = Stopwatch.StartNew();
                    RemoveDuplicateRows(DataTableBMES);
                    duplicateStopwatch.Stop();
                    clLogger.LogImportant($"Duplicate removal for new data completed in {duplicateStopwatch.Elapsed.TotalSeconds:F2}s");

                    var writeStopwatch = Stopwatch.StartNew();
                    sql.Writer.Write(CONSTANT.OPTION_TABLE_NAME.ORG, DataTableBMES);
                    writeStopwatch.Stop();
                    clLogger.LogImportant($"ORG write completed in {writeStopwatch.Elapsed.TotalSeconds:F2}s");
                }
                else
                {
                    // 기존 DB에 추가: 중복 제거 후 병합
                    clLogger.Log("Loading existing data from DB...");
                    var loadExistingStopwatch = Stopwatch.StartNew();
                    DataTable existingData = sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ORG);
                    loadExistingStopwatch.Stop();
                    clLogger.LogImportant($"Existing ORG load completed in {loadExistingStopwatch.Elapsed.TotalSeconds:F2}s");

                    if (existingData != null && existingData.Rows.Count > 0)
                    {
                        clLogger.Log($"Existing data rows: {existingData.Rows.Count}");
                        clLogger.Log($"New data rows: {DataTableBMES.Rows.Count}");

                        // 기존 데이터와 새 데이터 병합
                        var mergeStopwatch = Stopwatch.StartNew();
                        existingData.Merge(DataTableBMES);
                        mergeStopwatch.Stop();
                        clLogger.Log($"After merge: {existingData.Rows.Count} rows");
                        clLogger.LogImportant($"ORG merge completed in {mergeStopwatch.Elapsed.TotalSeconds:F2}s");

                        // 병합된 데이터에서 중복 행 제거 (같은 키를 가진 행 중 최신 것만 남김)
                        var duplicateStopwatch = Stopwatch.StartNew();
                        RemoveDuplicateRows(existingData);
                        duplicateStopwatch.Stop();

                        clLogger.Log($"After removing duplicates: {existingData.Rows.Count} rows");
                        clLogger.LogImportant($"Duplicate removal after merge completed in {duplicateStopwatch.Elapsed.TotalSeconds:F2}s");

                        // 기존 테이블 삭제 후 중복 제거된 데이터를 다시 쓰기
                        var rewriteStopwatch = Stopwatch.StartNew();
                        if (sql.Processor.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
                        {
                            sql.Processor.DropTable(CONSTANT.OPTION_TABLE_NAME.ORG);
                        }
                        sql.Writer.Write(CONSTANT.OPTION_TABLE_NAME.ORG, existingData);
                        rewriteStopwatch.Stop();
                        clLogger.LogImportant($"ORG rewrite completed in {rewriteStopwatch.Elapsed.TotalSeconds:F2}s");
                    }
                    else
                    {
                        // 기존 데이터가 없으면 새 데이터에서만 중복 제거
                        clLogger.Log("No existing data found. Removing duplicates from new data...");
                        var duplicateStopwatch = Stopwatch.StartNew();
                        RemoveDuplicateRows(DataTableBMES);
                        duplicateStopwatch.Stop();
                        clLogger.LogImportant($"Duplicate removal for new data completed in {duplicateStopwatch.Elapsed.TotalSeconds:F2}s");

                        var writeStopwatch = Stopwatch.StartNew();
                        sql.Writer.Write(CONSTANT.OPTION_TABLE_NAME.ORG, DataTableBMES);
                        writeStopwatch.Stop();
                        clLogger.LogImportant($"ORG write completed in {writeStopwatch.Elapsed.TotalSeconds:F2}s");
                    }
                }

                bool isRouting = sql.Processor.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING);
                bool isReason = sql.Processor.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON);
                
                if (!isRouting)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "txt Files|*.txt",
                        Title = "Open Routing file"
                    };

                    if (openFileDialog.ShowDialog() == true)
                    {
                        try
                        {
                            clLoadTxtFile ff = new clLoadTxtFile();

                            DataTable _d = await ff.Load(openFileDialog.FileName,
                                new string[] { CONSTANT.MATERIALNAME.NEW, CONSTANT.PROCESSCODE.NEW, CONSTANT.PROCESSNAME.NEW, CONSTANT.PROCESSTYPE.NEW }, '\t');

                            sql.Writer.Write(CONSTANT.OPTION_TABLE_NAME.ROUTING, _d);

                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                if (!isReason)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "txt Files|*.txt",
                        Title = "Open Reason file"
                    };

                    if (openFileDialog.ShowDialog() == true)
                    {
                        try
                        {
                            clLoadTxtFile ff = new clLoadTxtFile();

                            DataTable _d = await ff.Load(openFileDialog.FileName,
                                new string[] { CONSTANT.PROCESSNAME.NEW, CONSTANT.NGNAME.NEW, CONSTANT.REASON.NEW }, '\t');

                            sql.Writer.Write(CONSTANT.OPTION_TABLE_NAME.REASON, _d);

                        }
                        catch (Exception ex)
                        {

                        }
                    }

                }

                var lineShiftStopwatch = Stopwatch.StartNew();
                sql.Processor.SetLineShiftColumnValue(CONSTANT.OPTION_TABLE_NAME.ORG);
                lineShiftStopwatch.Stop();
                clLogger.LogImportant($"SetLineShiftColumnValue completed in {lineShiftStopwatch.Elapsed.TotalSeconds:F2}s");

                SaveOriginalTableUpdatedAt(dbFilePath, DateTime.Now);
                totalStopwatch.Stop();
                clLogger.Log("Successfully Save Data From BMES.");
                clLogger.LogImportant($"SaveBMESDataToDB completed in {totalStopwatch.Elapsed.TotalSeconds:F2}s");
                return dbFilePath; // Return DB path for auto-load
            }
        }

        private static void SaveOriginalTableUpdatedAt(string dbFilePath, DateTime updatedAt)
        {
            using var connection = new SQLiteConnection($"Data Source={dbFilePath};Version=3;");
            connection.Open();

            using var createCommand = new SQLiteCommand(
                $"CREATE TABLE IF NOT EXISTS [{OriginalTableMetaTableName}] (" +
                "[MetaKey] TEXT PRIMARY KEY, " +
                "[MetaValue] TEXT NOT NULL)",
                connection);
            createCommand.ExecuteNonQuery();

            using var upsertCommand = new SQLiteCommand(
                $"INSERT INTO [{OriginalTableMetaTableName}] ([MetaKey], [MetaValue]) VALUES (@key, @value) " +
                "ON CONFLICT([MetaKey]) DO UPDATE SET [MetaValue] = excluded.[MetaValue];",
                connection);
            upsertCommand.Parameters.AddWithValue("@key", OriginalTableUpdatedAtKey);
            upsertCommand.Parameters.AddWithValue("@value", updatedAt.ToString("yyyy-MM-dd HH:mm:ss"));
            upsertCommand.ExecuteNonQuery();
        }

        /// <summary>
        /// 테이블에서 중복 행 제거 (같은 키를 가진 행 중 최신 데이터만 남김)
        /// 키: PRODUCTION_LINE + PROCESSCODE + PROCESSNAME + NGNAME + NGCODE + MATERIALNAME + PRODUCT_DATE + SHIFT
        /// 최신 = 나중에 추가된 행 (인덱스가 큰 행)
        /// </summary>
        private void RemoveDuplicateRows(DataTable data)
        {
            if (data == null) return;

            var stopwatch = Stopwatch.StartNew();

            // 필수 컬럼 확인
            if (!data.Columns.Contains(CONSTANT.PRODUCTION_LINE.NEW) ||
                !data.Columns.Contains(CONSTANT.PROCESSCODE.NEW) ||
                !data.Columns.Contains(CONSTANT.PROCESSNAME.NEW) ||
                !data.Columns.Contains(CONSTANT.NGNAME.NEW) ||
                !data.Columns.Contains(CONSTANT.NGCODE.NEW) ||
                !data.Columns.Contains(CONSTANT.PRODUCT_DATE.NEW) ||
                !data.Columns.Contains(CONSTANT.MATERIALNAME.NEW) ||
                !data.Columns.Contains(CONSTANT.SHIFT.NEW))
            {
                clLogger.LogWarning("Cannot remove duplicates: table missing required columns");
                return;
            }

            clLogger.Log($"Starting duplicate removal. Total rows before: {data.Rows.Count}");

            // 각 키별로 마지막 행의 인덱스를 저장 (최신 데이터)
            var keyToLastIndex = new Dictionary<string, int>();
            var keyCounts = new Dictionary<string, int>();
            int duplicateCount = 0;

            for (int i = 0; i < data.Rows.Count; i++)
            {
                var row = data.Rows[i];

                // 키 기준:
                // PRODUCTION_LINE, PROCESSCODE, PROCESSNAME, NGNAME, NGCODE, MATERIALNAME, PRODUCT_DATE, SHIFT
                // QTYINPUT/QTYNG는 서버 재집계로 달라질 수 있으므로 키에서 제외
                string productionLine = row[CONSTANT.PRODUCTION_LINE.NEW]?.ToString()?.Trim() ?? "";
                string processCode = row[CONSTANT.PROCESSCODE.NEW]?.ToString()?.Trim() ?? "";
                string processName = CONSTANT.Normalize(row[CONSTANT.PROCESSNAME.NEW]?.ToString() ?? "");
                string ngName = CONSTANT.Normalize(row[CONSTANT.NGNAME.NEW]?.ToString() ?? "");
                string ngCode = row[CONSTANT.NGCODE.NEW]?.ToString()?.Trim() ?? "";
                string materialName = row[CONSTANT.MATERIALNAME.NEW]?.ToString()?.Trim() ?? "";
                string productDate = row[CONSTANT.PRODUCT_DATE.NEW]?.ToString()?.Trim() ?? "";
                string shift = row[CONSTANT.SHIFT.NEW]?.ToString()?.Trim() ?? "";

                string key = $"{productionLine}|{processCode}|{processName}|{ngName}|{ngCode}|{materialName}|{productDate}|{shift}";

                // 이미 존재하는 키면 중복
                if (keyToLastIndex.ContainsKey(key))
                {
                    duplicateCount++;
                }

                // 덮어쓰기 -> 마지막 인덱스가 남음 (최신 데이터)
                keyToLastIndex[key] = i;
                if (keyCounts.ContainsKey(key))
                {
                    keyCounts[key]++;
                }
                else
                {
                    keyCounts[key] = 1;
                }
            }

            int duplicateGroupCount = keyCounts.Count(kvp => kvp.Value > 1);
            clLogger.Log($"Found {duplicateCount} duplicate rows across {duplicateGroupCount} duplicate keys");

            // 유지할 인덱스 집합
            var rowsToKeep = new HashSet<int>(keyToLastIndex.Values);

            // 뒤에서부터 삭제 (인덱스 변화 방지)
            int removedCount = 0;
            for (int i = data.Rows.Count - 1; i >= 0; i--)
            {
                if (!rowsToKeep.Contains(i))
                {
                    data.Rows.RemoveAt(i);
                    removedCount++;
                }
            }

            clLogger.Log($"Removed {removedCount} duplicate rows (kept latest)");
            clLogger.Log($"Total rows after: {data.Rows.Count}");
            clLogger.Log($"Unique keys: {keyToLastIndex.Count}");
            stopwatch.Stop();
            clLogger.LogImportant($"RemoveDuplicateRows completed in {stopwatch.Elapsed.TotalSeconds:F2}s");
        }



    }
}
