using DataMaker.Logger;
using System.Data;
using System.Data.SQLite;

namespace DataMaker.R6.SQLProcess
{
    public interface ISQLiteDataProcessing
    {
        int GetRowCount(string tableName);
        void DropTable(string tableName);
        void CreateTable(string tableName, string[] columnNames, string[] columnTypes);
        void CreateTable(string tableName, string sourceTableName);
        List<string> GetListUniqueValuesOfColumns(string tableName, string columnName);
        void InsertData(string tableName, Dictionary<string, object> data);
        void UpdateData(string tableName, Dictionary<string, object> data, string condition);
        void DeleteData(string tableName, string condition);
        bool IsTableExist(string tableName);
        void CopyCommonColumns(string sourceTable, string targetTable);
        void CopyCommonColumns(string sourceTable, string targetTable, string columnName, List<string> values);
        void KeepRowsWhereInList(string tableName, List<string> columns, List<List<string>> values);
        void SetLineShiftColumnValue(string tableName);
        void SetEmptyColumnsValueInProcTable(string tableName);
        void NormalizeColumns(string tableName, params string[] columnNames);
        void UpdateTableFromTable(string targetTable, string sourceTable,
            (string TargetCol, string SourceCol)[] matchColumns,
            (string TargetCol, string SourceCol)[] updateColumns);
    }

    /// <summary>
    /// SQLite 데이터베이스의 복잡한 데이터 처리를 담당하는 클래스
    /// </summary>
    public class clSQLiteDataProcessing : BaseSqliteService, ISQLiteDataProcessing
    {
        #region Fields

        private readonly ISQLiteReader _reader;
        private readonly ISQLiteWriter _writer;

        #endregion

        #region Constructors

        public clSQLiteDataProcessing(string dbPath) : base(dbPath)
        {
            _reader = new clSQLiteReader(dbPath);
            _writer = new clSQLiteWriter(dbPath);
        }

        public clSQLiteDataProcessing(ISQLiteLoader loader, ISQLiteReader reader, ISQLiteWriter writer)
            : base(loader?.DBPath)
        {
            if (loader == null)
                throw new ArgumentNullException(nameof(loader), "Loader cannot be null.");

            _reader = reader ?? new clSQLiteReader(loader);
            _writer = writer ?? new clSQLiteWriter(loader.DBPath);
        }

        #endregion

        #region ISQLiteDataProcessing Implementation

        /// <summary>
        /// 테이블이 존재하는지 확인합니다.
        /// </summary>
        public bool IsTableExist(string tableName)
        {
            return TableExists(tableName);
        }

        /// <summary>
        /// 테이블의 행 개수를 반환합니다.
        /// </summary>
        public int GetRowCount(string tableName)
        {
            return base.GetRowCount(tableName);
        }

        /// <summary>
        /// 테이블을 삭제합니다.
        /// </summary>
        public void DropTable(string tableName)
        {
            if (!TableExists(tableName))
            {
                clLogger.Log($"Table '{tableName}' does not exist.");
                return;
            }

            string sql = $"DROP TABLE IF EXISTS [{tableName}]";
            ExecuteNonQuery(sql);
        }

        /// <summary>
        /// 테이블을 생성합니다.
        /// </summary>
        public void CreateTable(string tableName, string[] columnNames, string[] columnTypes)
        {
            _writer.CreateTable(tableName, columnNames, columnTypes);
        }

        /// <summary>
        /// 다른 테이블의 구조를 복사하여 새 테이블을 생성합니다.
        /// </summary>
        public void CreateTable(string tableName, string sourceTableName)
        {
            if (!TableExists(sourceTableName))
                throw new InvalidOperationException($"Source table '{sourceTableName}' does not exist.");

            string sql = $"CREATE TABLE IF NOT EXISTS [{tableName}] AS SELECT * FROM [{sourceTableName}] WHERE 0;";
            ExecuteNonQuery(sql);
        }

        /// <summary>
        /// 컬럼의 고유 값 목록을 반환합니다.
        /// </summary>
        public List<string> GetListUniqueValuesOfColumns(string tableName, string columnName)
        {
            var values = new List<string>();
            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                string sql = $"SELECT DISTINCT [{columnName}] FROM [{tableName}] WHERE [{columnName}] IS NOT NULL;";

                using (var cmd = new SQLiteCommand(sql, Connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        values.Add(reader[columnName]?.ToString());
                    }
                }
            }
            finally
            {
                if (!wasOpen)
                {
                    CloseConnection();
                }
            }

            return values;
        }

        /// <summary>
        /// 데이터를 삽입합니다.
        /// </summary>
        public void InsertData(string tableName, Dictionary<string, object> data)
        {
            if (data == null || data.Count == 0)
                throw new ArgumentException("Data cannot be null or empty.", nameof(data));

            var columns = string.Join(", ", data.Keys.Select(k => $"[{k}]"));
            var parameters = string.Join(", ", data.Keys.Select(k => $"@{k}"));

            string sql = $"INSERT INTO [{tableName}] ({columns}) VALUES ({parameters});";

            var sqlParams = data.ToDictionary(kv => $"@{kv.Key}", kv => kv.Value);
            ExecuteNonQuery(sql, sqlParams);
        }

        /// <summary>
        /// 조건에 맞는 데이터를 업데이트합니다.
        /// </summary>
        public void UpdateData(string tableName, Dictionary<string, object> data, string condition)
        {
            if (data == null || data.Count == 0)
                throw new ArgumentException("Data cannot be null or empty.", nameof(data));

            var setClause = string.Join(", ", data.Keys.Select(k => $"[{k}] = @{k}"));
            string sql = $"UPDATE [{tableName}] SET {setClause}";

            if (!string.IsNullOrWhiteSpace(condition))
            {
                sql += $" WHERE {condition}";
            }

            var sqlParams = data.ToDictionary(kv => $"@{kv.Key}", kv => kv.Value);
            ExecuteNonQuery(sql, sqlParams);
        }

        /// <summary>
        /// 조건에 맞는 데이터를 삭제합니다.
        /// </summary>
        public void DeleteData(string tableName, string condition)
        {
            string sql = $"DELETE FROM [{tableName}]";

            if (!string.IsNullOrWhiteSpace(condition))
            {
                sql += $" WHERE {condition}";
            }

            ExecuteNonQuery(sql);
        }

        /// <summary>
        /// 두 테이블의 공통 컬럼 데이터를 복사합니다.
        /// </summary>
        public void CopyCommonColumns(string sourceTable, string targetTable)
        {
            // 소스 및 타겟 테이블 컬럼 조회
            var sourceCols = _reader.GetColumns(sourceTable);
            var targetCols = _reader.GetColumns(targetTable);

            // 교집합 컬럼 추출
            var commonCols = sourceCols.Intersect(targetCols).ToList();

            if (commonCols.Count == 0)
                throw new InvalidOperationException("No common columns found between tables.");

            string colList = string.Join(", ", commonCols.Select(c => $"[{c}]"));
            string sql = $"INSERT INTO [{targetTable}] ({colList}) SELECT {colList} FROM [{sourceTable}];";

            ExecuteNonQuery(sql);
        }

        /// <summary>
        /// 특정 값들에 해당하는 행만 복사합니다.
        /// </summary>
        public void CopyCommonColumns(string sourceTable, string targetTable, string columnName, List<string> values)
        {
            if (values == null || values.Count == 0)
                return;

            var sourceCols = _reader.GetColumns(sourceTable);
            var targetCols = _reader.GetColumns(targetTable);
            var commonCols = sourceCols.Intersect(targetCols).ToList();

            if (commonCols.Count == 0)
                throw new InvalidOperationException("No common columns found between tables.");

            string colList = string.Join(", ", commonCols.Select(c => $"[{c}]"));
            var parameters = string.Join(", ", values.Select((v, i) => $"@p{i}"));
            string sql = $@"
                INSERT INTO [{targetTable}] ({colList})
                SELECT {colList}
                FROM [{sourceTable}]
                WHERE [{columnName}] IN ({parameters});";

            var sqlParams = new Dictionary<string, object>();
            for (int i = 0; i < values.Count; i++)
            {
                sqlParams.Add($"@p{i}", values[i]);
            }

            ExecuteNonQuery(sql, sqlParams);
        }

        /// <summary>
        /// 조건에 맞는 행만 유지합니다.
        /// </summary>
        public void KeepRowsWhereInList(string tableName, List<string> columns, List<List<string>> values)
        {
            if (columns == null || columns.Count == 0)
                throw new ArgumentException("Columns cannot be null or empty.", nameof(columns));

            if (values == null || values.Count == 0)
                throw new ArgumentException("Values cannot be null or empty.", nameof(values));

            // 임시 테이블 생성 및 필터링 로직 구현
            // (복잡한 로직이므로 필요에 따라 구현)
            throw new NotImplementedException("This method needs specific implementation based on requirements.");
        }

        /// <summary>
        /// 모델 필터링에 필요한 LineShift 컬럼만 먼저 채웁니다.
        /// </summary>
        public void SetLineShiftColumnValue(string tableName)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            if (!TableExists(tableName))
                throw new InvalidOperationException($"Table '{tableName}' does not exist.");

            var columns = _reader.GetColumns(tableName);
            if (!columns.Contains(CONSTANT.ModelName_WithLineShift.NEW))
            {
                string alterSql = $"ALTER TABLE [{tableName}] ADD COLUMN [{CONSTANT.ModelName_WithLineShift.NEW}] TEXT;";
                ExecuteNonQuery(alterSql);
            }

            string updateSql = $@"
                UPDATE [{tableName}]
                SET
                    [{CONSTANT.ModelName_WithLineShift.NEW}] = MATERIALNAME || '_' || PRODUCTION_LINE";

            ExecuteNonQuery(updateSql);
        }

        /// <summary>
        /// ProcTable의 파생 컬럼들을 계산하여 채웁니다.
        /// MONTH, WEEK, LINE_REMOVE, LINE, LINESHIFT, LR, LRLINE, LR_BUILDING 등의 값을 설정합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        public void SetEmptyColumnsValueInProcTable(string tableName)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            // 파생 컬럼 값들을 계산하여 업데이트
            string updateSql = $@"
                UPDATE [{tableName}]
                SET
                    MONTH       = CAST(strftime('%m', DATE(PRODUCT_DATE)) AS INTEGER),
                    WEEK        = CAST(strftime('%W', DATE(PRODUCT_DATE)) AS INTEGER) + 1,
                    LINESHIFT = MATERIALNAME || '_' || PRODUCTION_LINE";
                    /*LINE_REMOVE = substr(PRODUCTION_LINE, 1, 3),
                    LINE        = MATERIALNAME || '_' || substr(PRODUCTION_LINE, 1, 3),
                    LINESHIFT   = MATERIALNAME || '_' || PRODUCTION_LINE,
                    LR          = REPLACE(REPLACE(MATERIALNAME, '-L-', '-'), '-R-', '-'),
                    LRLINE      = REPLACE(REPLACE(MATERIALNAME, '-L-', '-'), '-R-', '-') 
                                  || '_' || substr(PRODUCTION_LINE, 1, 3),
                    LR_BUILDING = REPLACE(REPLACE(MATERIALNAME, '-L-', '-'), '-R-', '-') 
                                  || '_' || substr(PRODUCTION_LINE, 1, 2);";*/

            ExecuteNonQuery(updateSql);
        }

        /// <summary>
        /// 지정된 컬럼들의 값을 정규화합니다.
        /// </summary>
        public void NormalizeColumns(string tableName, params string[] columnNames)
        {
            if (columnNames == null || columnNames.Length == 0)
                throw new ArgumentException("Column names cannot be null or empty.", nameof(columnNames));

            ExecuteInTransaction(transaction =>
            {
                string colList = string.Join(", ", columnNames.Select(c => $"[{c}]"));
                string selectSql = $"SELECT rowid, {colList} FROM [{tableName}];";

                using (var selectCmd = new SQLiteCommand(selectSql, Connection, transaction))
                using (var reader = selectCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        long rowId = reader.GetInt64(0);

                        var newValues = new Dictionary<string, string>();
                        foreach (var col in columnNames)
                        {
                            string original = reader[col]?.ToString() ?? "";
                            string normalized = NormalizeString(original);
                            newValues[col] = normalized;
                        }

                        string setClause = string.Join(", ", newValues.Keys.Select(c => $"[{c}] = @{c}"));
                        string updateSql = $"UPDATE [{tableName}] SET {setClause} WHERE rowid = @rowid;";

                        using (var updateCmd = new SQLiteCommand(updateSql, Connection, transaction))
                        {
                            foreach (var kv in newValues)
                                updateCmd.Parameters.AddWithValue($"@{kv.Key}", kv.Value);

                            updateCmd.Parameters.AddWithValue("@rowid", rowId);
                            updateCmd.ExecuteNonQuery();
                        }
                    }
                }
            });
        }

        /// <summary>
        /// 다른 테이블의 데이터로 현재 테이블을 업데이트합니다.
        /// </summary>
        public void UpdateTableFromTable(
            string targetTable,
            string sourceTable,
            (string TargetCol, string SourceCol)[] matchColumns,
            (string TargetCol, string SourceCol)[] updateColumns)
        {
            if (matchColumns == null || matchColumns.Length == 0)
                throw new ArgumentException("Match columns cannot be null or empty.", nameof(matchColumns));

            if (updateColumns == null || updateColumns.Length == 0)
                throw new ArgumentException("Update columns cannot be null or empty.", nameof(updateColumns));

            ExecuteInTransaction(transaction =>
            {
                UpdateWithTempTableAndIndex(targetTable, sourceTable, matchColumns, updateColumns, transaction);
            });
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// 임시 테이블과 인덱스를 사용한 업데이트
        /// </summary>
        private void UpdateWithTempTableAndIndex(
            string targetTable,
            string sourceTable,
            (string TargetCol, string SourceCol)[] matchColumns,
            (string TargetCol, string SourceCol)[] updateColumns,
            SQLiteTransaction transaction)
        {
            string tempTableName = $"temp_update_{Guid.NewGuid():N}";

            try
            {
                // 매칭 키 생성
                string matchKeyExpression = string.Join(" || '|' || ",
                    matchColumns.Select(m => $"COALESCE(CAST([{m.SourceCol}] AS TEXT), '')"));

                string targetMatchKeyExpression = string.Join(" || '|' || ",
                    matchColumns.Select(m => $"COALESCE(CAST([{m.TargetCol}] AS TEXT), '')"));

                // 임시 테이블 생성
                string tempTableColumns = "match_key TEXT, " +
                    string.Join(", ", updateColumns.Select(u => $"[{u.SourceCol}] TEXT"));

                using (var cmd = new SQLiteCommand($"CREATE TEMP TABLE [{tempTableName}] ({tempTableColumns})",
                    Connection, transaction))
                {
                    cmd.ExecuteNonQuery();
                }

                // 인덱스 생성
                using (var cmd = new SQLiteCommand($"CREATE INDEX idx_{tempTableName}_match ON [{tempTableName}] (match_key)",
                    Connection, transaction))
                {
                    cmd.ExecuteNonQuery();
                }

                // 소스 데이터 삽입
                string insertColumns = "match_key, " + string.Join(", ", updateColumns.Select(u => $"[{u.SourceCol}]"));
                string selectColumns = $"({matchKeyExpression}) as match_key, " +
                    string.Join(", ", updateColumns.Select(u => $"[{u.SourceCol}]"));

                using (var cmd = new SQLiteCommand($@"
                    INSERT INTO [{tempTableName}] ({insertColumns})
                    SELECT {selectColumns} FROM [{sourceTable}]",
                    Connection, transaction))
                {
                    cmd.ExecuteNonQuery();
                }

                // 배치 업데이트 실행
                foreach (var updateCol in updateColumns)
                {
                    using (var cmd = new SQLiteCommand($@"
                        UPDATE [{targetTable}] 
                        SET [{updateCol.TargetCol}] = (
                            SELECT [{updateCol.SourceCol}] 
                            FROM [{tempTableName}] 
                            WHERE match_key = ({targetMatchKeyExpression})
                        )
                        WHERE EXISTS (
                            SELECT 1 FROM [{tempTableName}] 
                            WHERE match_key = ({targetMatchKeyExpression})
                        )", Connection, transaction))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                // 임시 테이블 정리
                try
                {
                    using (var cmd = new SQLiteCommand($"DROP TABLE IF EXISTS [{tempTableName}]",
                        Connection, transaction))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch { /* 무시 */ }
            }
        }

        /// <summary>
        /// 문자열을 정규화합니다.
        /// </summary>
        private string NormalizeString(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return input;

            // 기본 정규화 로직 (필요에 따라 커스터마이징)
            return input.Trim()
                       .Replace("  ", " ")  // 연속 공백 제거
                       .ToUpperInvariant(); // 대문자 변환
        }

        #endregion
    }
}
