using System.Data;
using System.Data.SQLite;

namespace DataMaker.R6.SQLProcess
{
    public interface ISQLiteWriter
    {
        void Write(string tableName, DataTable table);
        void Append(string tableName, DataTable table);
        void CreateTable(string tableName, string[] columnNames, string[] columnTypes);
    }

    /// <summary>
    /// SQLite 데이터베이스에 데이터를 쓰는 클래스
    /// </summary>
    public class clSQLiteWriter : BaseSqliteService, ISQLiteWriter
    {
        #region Constructors

        public clSQLiteWriter() : base()
        {
        }

        public clSQLiteWriter(string dbPath) : base(dbPath)
        {
        }

        /// <summary>
        /// Loader를 사용하는 생성자 (기존 호환성 유지)
        /// </summary>
        /// <param name="loader">SQLite 로더</param>
        public clSQLiteWriter(clSQLiteLoader loader) : base(loader?.DBPath)
        {
            if (loader == null)
                throw new ArgumentNullException(nameof(loader), "Loader cannot be null.");
        }

        #endregion

        #region ISQLiteWriter Implementation

        /// <summary>
        /// 테이블을 생성하고 데이터를 씁니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="table">쓸 데이터</param>
        public void Write(string tableName, DataTable table)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            if (table == null)
                throw new ArgumentNullException(nameof(table), "DataTable cannot be null.");

            // 컬럼 정보 추출
            string[] columnNames = table.Columns
                                .Cast<DataColumn>()
                                .Select(c => c.ColumnName)
                                .ToArray();

            string[] columnTypes = table.Columns
                              .Cast<DataColumn>()
                              .Select(c => GetSQLiteType(c.DataType))
                              .ToArray();

            // 테이블 생성
            CreateTable(tableName, columnNames, columnTypes);

            // 데이터 삽입
            Append(tableName, table);
        }

        /// <summary>
        /// 기존 테이블에 데이터를 추가합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="table">추가할 데이터</param>
        public void Append(string tableName, DataTable table)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            if (table == null || table.Rows.Count == 0)
                return;

            InsertDataTable(table, tableName);
        }

        /// <summary>
        /// 테이블을 생성합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="columnNames">컬럼 이름 배열</param>
        /// <param name="columnTypes">컬럼 타입 배열</param>
        public void CreateTable(string tableName, string[] columnNames, string[] columnTypes)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            if (columnNames == null || columnNames.Length == 0)
                throw new ArgumentException("Column names cannot be null or empty.", nameof(columnNames));

            if (columnTypes == null || columnTypes.Length == 0)
                throw new ArgumentException("Column types cannot be null or empty.", nameof(columnTypes));

            if (columnNames.Length != columnTypes.Length)
                throw new ArgumentException("Column names and types must have the same length.");

            // 컬럼 정의 생성
            var columnDefinitions = new System.Text.StringBuilder();
            for (int i = 0; i < columnNames.Length; i++)
            {
                columnDefinitions.Append($"[{columnNames[i]}] {columnTypes[i]}");
                if (i < columnNames.Length - 1)
                    columnDefinitions.Append(", ");
            }

            string sql = $"CREATE TABLE IF NOT EXISTS [{tableName}] ({columnDefinitions});";
            ExecuteNonQuery(sql);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// DataTable의 데이터를 삽입합니다.
        /// </summary>
        /// <param name="table">삽입할 DataTable</param>
        /// <param name="tableName">대상 테이블 이름</param>
        private void InsertDataTable(DataTable table, string tableName)
        {
            if (table.Rows.Count == 0)
                return;

            ExecuteInTransaction(tran =>
            {
                InsertWithBatchValues(table, tableName, tran);
            });
        }

        /// <summary>
        /// 배치 INSERT를 수행합니다.
        /// </summary>
        /// <param name="table">삽입할 DataTable</param>
        /// <param name="tableName">대상 테이블 이름</param>
        /// <param name="transaction">트랜잭션</param>
        private void InsertWithBatchValues(DataTable table, string tableName, SQLiteTransaction transaction)
        {
            // INSERT SQL 준비
            var columnNames = string.Join(", ", table.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}]"));
            var paramPlaceholder = "(" + string.Join(", ", table.Columns.Cast<DataColumn>().Select((c, idx) => $"@p{idx}")) + ")";
            string sql = $"INSERT INTO [{tableName}] ({columnNames}) VALUES {paramPlaceholder};";

            using (var cmd = new SQLiteCommand(sql, Connection, transaction))
            {
                // 파라미터 미리 생성
                foreach (DataColumn col in table.Columns)
                {
                    cmd.Parameters.Add(new SQLiteParameter($"@p{col.Ordinal}"));
                }

                // 각 행 삽입
                foreach (DataRow row in table.Rows)
                {
                    for (int colIdx = 0; colIdx < table.Columns.Count; colIdx++)
                    {
                        cmd.Parameters[colIdx].Value = row[colIdx] ?? DBNull.Value;
                    }
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// .NET Type을 SQLite Type으로 변환합니다.
        /// </summary>
        /// <param name="type">.NET 타입</param>
        /// <returns>SQLite 타입 문자열</returns>
        private string GetSQLiteType(Type type)
        {
            if (type == typeof(int) || type == typeof(long) || type == typeof(short) ||
                type == typeof(byte) || type == typeof(bool))
                return "INTEGER";

            if (type == typeof(float) || type == typeof(double) || type == typeof(decimal))
                return "REAL";

            if (type == typeof(DateTime))
                return "TEXT"; // SQLite는 날짜를 TEXT나 INTEGER로 저장

            if (type == typeof(byte[]))
                return "BLOB";

            // 기본값은 TEXT
            return "TEXT";
        }

        #endregion
    }
}
