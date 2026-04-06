using DataMaker.Logger;
using DataMaker.R6;
using DataMaker.R6.PreProcessor;
using DataMaker.R6.SQLProcess;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DataMaker.R6
{
    /// <summary>
    /// 데이터 처리 워크플로우 관리 클래스
    /// Button 클릭 → DB 로드 → CSV 로드 → OriginalTable 준비
    /// </summary>
    public class clDataProcessor
    {
        private readonly string _dbPath;
        private readonly string _processTypeCsvPath;
        private readonly string _reasonCsvPath;
        private clSQLFileIO _sql;
        private readonly Action<int, string>? _progressCallback;

        public clDataProcessor(string dbPath, string processTypeCsvPath, string reasonCsvPath, Action<int, string>? progressCallback = null)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
                throw new ArgumentException("Database path cannot be null or empty.", nameof(dbPath));

            if (string.IsNullOrWhiteSpace(processTypeCsvPath))
                throw new ArgumentException("ProcessType CSV path cannot be null or empty.", nameof(processTypeCsvPath));

            if (string.IsNullOrWhiteSpace(reasonCsvPath))
                throw new ArgumentException("Reason CSV path cannot be null or empty.", nameof(reasonCsvPath));

            _dbPath = dbPath;
            _processTypeCsvPath = processTypeCsvPath;
            _reasonCsvPath = reasonCsvPath;
            _progressCallback = progressCallback;
        }

        /// <summary>
        /// 전체 프로세스 실행: DB 로드 → CSV 로드 → OriginalTable 준비
        /// </summary>
        public void ProcessData()
        {
            try
            {
                clLogger.Log("=== Data Processing Started ===");
                ReportProgress(5, "Validating DB and initializing connection...");

                // 1. DB 연결 초기화
                InitializeDatabase();
                ReportProgress(20, "Database verified. Loading Routing/Reason tables...");

                // 2. CSV 파일 로드 (ProcessTypeTable, ReasonTable)
                LoadCsvTables();
                ReportProgress(75, "CSV tables loaded. Preparing OriginalTable...");

                // 3. OriginalTable을 ProcTable 생성 가능 상태로 준비
                PrepareOriginalTable();
                ReportProgress(100, "OriginalTable preparation completed.");

                clLogger.Log("=== Data Processing Completed Successfully ===");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error in ProcessData");
                throw;
            }
        }

        /// <summary>
        /// 1단계: 데이터베이스 초기화
        /// </summary>
        private void InitializeDatabase()
        {
            clLogger.Log("Step 1: Initializing Database Connection");
            ReportProgress(10, "Checking selected DB file...");

            if (!File.Exists(_dbPath))
            {
                throw new FileNotFoundException($"Database file not found: {_dbPath}");
            }

            _sql = new clSQLFileIO(_dbPath);

            // OriginalTable 존재 확인
            if (!_sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
            {
                throw new InvalidOperationException($"Table '{CONSTANT.OPTION_TABLE_NAME.ORG}' does not exist in database.");
            }

            int rowCount = _sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.ORG);
            clLogger.Log($"  - OriginalTable loaded: {rowCount} rows");
            ReportProgress(20, $"OriginalTable verified: {rowCount:N0} rows");
        }

        /// <summary>
        /// 2단계: CSV 파일을 DB 테이블로 로드
        /// </summary>
        private void LoadCsvTables()
        {
            clLogger.Log("Step 2: Loading CSV Files to Database");
            ReportProgress(25, "Loading Routing table from CSV...");

            // ProcessTypeTable 로드
            LoadProcessTypeTable();
            ReportProgress(50, "Routing table completed. Loading Reason table...");

            // ReasonTable 로드
            LoadReasonTable();
            ReportProgress(75, "Routing/Reason tables completed.");
        }

        /// <summary>
        /// ProcessTypeTable CSV → DB
        /// </summary>
        private void LoadProcessTypeTable()
        {
            clLogger.Log("  - Loading ProcessTypeTable from CSV");
            ReportProgress(30, "Loading ProcessTypeTable...");

            if (!File.Exists(_processTypeCsvPath))
            {
                throw new FileNotFoundException($"ProcessType CSV file not found: {_processTypeCsvPath}");
            }

            // 기존 테이블 삭제
            if (_sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING))
            {
                _sql.DropTable(CONSTANT.OPTION_TABLE_NAME.ROUTING);
                clLogger.Log("    - Existing ProcessTypeTable dropped");
            }

            // CSV → DataTable → DB
            var txtTableMaker = new clMakeTxtTable(_dbPath);
            txtTableMaker.Make(
                CONSTANT.OPTION_TABLE_NAME.ROUTING,
                _processTypeCsvPath,
                clOption.GetProcessTypeTableColumns()
            );

            int rowCount = _sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.ROUTING);
            clLogger.Log($"    - ProcessTypeTable created: {rowCount} rows");
            ReportProgress(40, $"ProcessTypeTable created: {rowCount:N0} rows");

            // ProcessName, ProcessType에 CONSTANT.Normalize 적용
            NormalizeProcessTypeTable();
            ReportProgress(50, "ProcessTypeTable normalization completed.");
        }

        /// <summary>
        /// ProcessTypeTable의 ProcessName, ProcessType 컬럼 정규화
        /// </summary>
        private void NormalizeProcessTypeTable()
        {
            clLogger.Log("    - Normalizing ProcessTypeTable (ProcessName, ProcessType) with CONSTANT.Normalize");

            var routingData = _sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ROUTING);

            if (routingData == null || routingData.Rows.Count == 0)
            {
                clLogger.LogWarning("    - ProcessTypeTable is empty, skipping normalization");
                return;
            }

            foreach (System.Data.DataRow row in routingData.Rows)
            {
                if (routingData.Columns.Contains("ProcessName"))
                {
                    string original = row["ProcessName"]?.ToString() ?? "";
                    row["ProcessName"] = CONSTANT.Normalize(original);
                }

                if (routingData.Columns.Contains("ProcessType"))
                {
                    string original = row["ProcessType"]?.ToString() ?? "";
                    row["ProcessType"] = CONSTANT.Normalize(original);
                }
            }

            _sql.DropTable(CONSTANT.OPTION_TABLE_NAME.ROUTING);
            _sql.CreateTable(
                CONSTANT.OPTION_TABLE_NAME.ROUTING,
                clOption.GetProcessTypeTableColumns().ToArray(),
                new[] { "TEXT", "TEXT", "TEXT", "TEXT" }
            );
            _sql.Writer.Write(CONSTANT.OPTION_TABLE_NAME.ROUTING, routingData);

            clLogger.Log("    - ProcessTypeTable normalization completed");
        }

        /// <summary>
        /// ReasonTable CSV → DB
        /// </summary>
        private void LoadReasonTable()
        {
            clLogger.Log("  - Loading ReasonTable from CSV");
            ReportProgress(55, "Loading ReasonTable...");

            if (!File.Exists(_reasonCsvPath))
            {
                throw new FileNotFoundException($"Reason CSV file not found: {_reasonCsvPath}");
            }

            // 기존 테이블 삭제
            if (_sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON))
            {
                _sql.DropTable(CONSTANT.OPTION_TABLE_NAME.REASON);
                clLogger.Log("    - Existing ReasonTable dropped");
            }

            // CSV → DataTable → DB
            var txtTableMaker = new clMakeTxtTable(_dbPath);
            txtTableMaker.Make(
                CONSTANT.OPTION_TABLE_NAME.REASON,
                _reasonCsvPath,
                clOption.GetReasonTableColumns()
            );

            int rowCount = _sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.REASON);
            clLogger.Log($"    - ReasonTable created: {rowCount} rows");
            ReportProgress(65, $"ReasonTable created: {rowCount:N0} rows");

            // Normalize 적용 (processName, NgName) - AccessTable과 동일한 방식
            NormalizeReasonTable();
            ReportProgress(75, "ReasonTable normalization completed.");
        }

        /// <summary>
        /// ReasonTable의 processName, NgName 컬럼 정규화
        /// AccessTable과 동일한 CONSTANT.Normalize() 사용
        /// </summary>
        private void NormalizeReasonTable()
        {
            clLogger.Log("    - Normalizing ReasonTable (processName, NgName) with CONSTANT.Normalize");

            // DataTable로 로드
            var reasonData = _sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.REASON);

            if (reasonData == null || reasonData.Rows.Count == 0)
            {
                clLogger.LogWarning("    - ReasonTable is empty, skipping normalization");
                return;
            }

            // processName, NgName에 CONSTANT.Normalize 적용
            foreach (System.Data.DataRow row in reasonData.Rows)
            {
                if (reasonData.Columns.Contains("processName"))
                {
                    string original = row["processName"]?.ToString() ?? "";
                    row["processName"] = CONSTANT.Normalize(original);
                }

                if (reasonData.Columns.Contains("NgName"))
                {
                    string original = row["NgName"]?.ToString() ?? "";
                    row["NgName"] = CONSTANT.Normalize(original);
                }
            }

            // 테이블 재작성
            _sql.DropTable(CONSTANT.OPTION_TABLE_NAME.REASON);
            _sql.CreateTable(
                CONSTANT.OPTION_TABLE_NAME.REASON,
                clOption.GetReasonTableColumns().ToArray(),
                new[] { "TEXT", "TEXT", "TEXT" }
            );
            _sql.Writer.Write(CONSTANT.OPTION_TABLE_NAME.REASON, reasonData);

            clLogger.Log("    - ReasonTable normalization completed (using CONSTANT.Normalize)");
        }

        /// <summary>
        /// 3단계: OriginalTable 준비
        /// </summary>
        private void PrepareOriginalTable()
        {
            clLogger.Log("Step 3: Preparing OriginalTable for model filtering");
            ReportProgress(85, "Preparing OriginalTable for model filtering...");
            _sql.Processor.SetLineShiftColumnValue(CONSTANT.OPTION_TABLE_NAME.ORG);
            int rowCount = _sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.ORG);
            clLogger.Log($"  - OriginalTable prepared: {rowCount} rows with LineShift");
            ReportProgress(100, $"OriginalTable prepared: {rowCount:N0} rows");
        }

        private void ReportProgress(int percent, string message)
        {
            _progressCallback?.Invoke(percent, message);
        }

        /// <summary>
        /// ProcessTypeTable만 업데이트 (외부에서 호출 가능)
        /// </summary>
        public void UpdateProcessTypeTable(string dbPath, string processTypeCsvPath)
        {
            try
            {
                clLogger.Log("=== Updating ProcessTypeTable ===");

                if (_sql == null || _dbPath != dbPath)
                {
                    _sql?.Dispose();
                    _sql = new clSQLFileIO(dbPath);
                }

                LoadProcessTypeTableFromFile(processTypeCsvPath);

                clLogger.Log("=== ProcessTypeTable Update Completed Successfully ===");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error in UpdateProcessTypeTable");
                throw;
            }
        }

        /// <summary>
        /// 파일에서 ProcessTypeTable 로드 (public 메서드용)
        /// </summary>
        private void LoadProcessTypeTableFromFile(string processTypeCsvPath)
        {
            clLogger.Log("  - Loading ProcessTypeTable from CSV");

            if (!File.Exists(processTypeCsvPath))
            {
                throw new FileNotFoundException($"ProcessType CSV file not found: {processTypeCsvPath}");
            }

            if (_sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING))
            {
                _sql.DropTable(CONSTANT.OPTION_TABLE_NAME.ROUTING);
                clLogger.Log("    - Existing ProcessTypeTable dropped");
            }

            var txtTableMaker = new clMakeTxtTable(_sql.Processor.DBPath);
            txtTableMaker.Make(
                CONSTANT.OPTION_TABLE_NAME.ROUTING,
                processTypeCsvPath,
                clOption.GetProcessTypeTableColumns()
            );

            int rowCount = _sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.ROUTING);
            clLogger.Log($"    - ProcessTypeTable created: {rowCount} rows");

            NormalizeProcessTypeTable();
        }

        /// <summary>
        /// ReasonTable만 업데이트 (외부에서 호출 가능)
        /// </summary>
        public void UpdateReasonTable(string dbPath, string reasonCsvPath)
        {
            try
            {
                clLogger.Log("=== Updating ReasonTable ===");

                // DB 초기화
                if (_sql == null || _dbPath != dbPath)
                {
                    _sql?.Dispose();
                    _sql = new clSQLFileIO(dbPath);
                }

                // ReasonTable 로드
                LoadReasonTableFromFile(reasonCsvPath);

                clLogger.Log("=== ReasonTable Update Completed Successfully ===");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error in UpdateReasonTable");
                throw;
            }
        }

        /// <summary>
        /// 파일에서 ReasonTable 로드 (public 메서드용)
        /// </summary>
        private void LoadReasonTableFromFile(string reasonCsvPath)
        {
            clLogger.Log("  - Loading ReasonTable from CSV");

            if (!File.Exists(reasonCsvPath))
            {
                throw new FileNotFoundException($"Reason CSV file not found: {reasonCsvPath}");
            }

            // 기존 테이블 삭제
            if (_sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON))
            {
                _sql.DropTable(CONSTANT.OPTION_TABLE_NAME.REASON);
                clLogger.Log("    - Existing ReasonTable dropped");
            }

            // CSV → DataTable → DB
            var txtTableMaker = new clMakeTxtTable(_sql.Processor.DBPath);
            txtTableMaker.Make(
                CONSTANT.OPTION_TABLE_NAME.REASON,
                reasonCsvPath,
                clOption.GetReasonTableColumns()
            );

            int rowCount = _sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.REASON);
            clLogger.Log($"    - ReasonTable created: {rowCount} rows");

            // Normalize 적용 (processName, NgName) - AccessTable과 동일한 방식
            NormalizeReasonTable();
        }

        /// <summary>
        /// 리소스 해제
        /// </summary>
        public void Dispose()
        {
            _sql?.Dispose();
        }
    }
}
