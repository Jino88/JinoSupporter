using System.Data;
using System.Data.SQLite;

namespace DataMaker.R6.SQLProcess
{
    public interface ISQLiteLoader
    {
        string DBPath { get; set; }
        SQLiteConnection Connection { get; }
        void OpenConnection();
        void CloseConnection();
        DataTable Load(string tableName);
        DataTable Load(string tableName, string modelName, string columnName);
        DataTable Load(List<string> tableNames);
    }

    /// <summary>
    /// SQLite 데이터베이스에서 데이터를 로드하는 클래스
    /// </summary>
    public class clSQLiteLoader : BaseSqliteService, ISQLiteLoader
    {
        #region Constructors

        public clSQLiteLoader() : base()
        {
        }

        public clSQLiteLoader(string dbPath) : base(dbPath)
        {
        }

        #endregion

        #region ISQLiteLoader Implementation

        SQLiteConnection ISQLiteLoader.Connection => Connection;

        #endregion

        #region Load Methods

        /// <summary>
        /// 테이블 전체를 로드합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>로드된 DataTable</returns>
        public DataTable Load(string tableName)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            if (string.IsNullOrEmpty(DBPath))
                throw new InvalidOperationException("Database path is not set. Please set the DBPath property before loading the table.");

            return LoadInternal(tableName);
        }

        /// <summary>
        /// 조건에 맞는 데이터를 로드합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="modelName">모델 이름 (조건 값)</param>
        /// <param name="columnName">컬럼 이름 (조건 컬럼)</param>
        /// <returns>로드된 DataTable</returns>
        public DataTable Load(string tableName, string modelName, string columnName)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            if (string.IsNullOrEmpty(columnName))
                throw new ArgumentException("Column name cannot be null or empty.", nameof(columnName));

            if (string.IsNullOrEmpty(DBPath))
                throw new InvalidOperationException("Database path is not set. Please set the DBPath property before loading the table.");

            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                DataTable dataTable = new DataTable(tableName);

                string sql = $"SELECT * FROM [{tableName}] WHERE [{columnName}] = @ModelName";
                using (var cmd = new SQLiteCommand(sql, Connection))
                {
                    cmd.Parameters.AddWithValue("@ModelName", modelName);

                    using (var adapter = new SQLiteDataAdapter(cmd))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                return dataTable;
            }
            finally
            {
                if (!wasOpen)
                {
                    CloseConnection();
                }
            }
        }

        /// <summary>
        /// 여러 테이블을 로드하여 병합합니다.
        /// </summary>
        /// <param name="tableNames">테이블 이름 목록</param>
        /// <returns>병합된 DataTable</returns>
        public DataTable Load(List<string> tableNames)
        {
            if (tableNames == null || tableNames.Count == 0)
                throw new ArgumentException("Table names cannot be null or empty.", nameof(tableNames));

            DataTable combinedDataTable = new DataTable();

            foreach (var tableName in tableNames)
            {
                var dataTable = Load(tableName);

                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    if (combinedDataTable.Columns.Count == 0)
                    {
                        // 첫 번째 테이블의 스키마를 복사
                        combinedDataTable = dataTable.Copy();
                    }
                    else
                    {
                        // 이후 테이블의 행들을 추가
                        foreach (DataRow row in dataTable.Rows)
                        {
                            combinedDataTable.ImportRow(row);
                        }
                    }
                }
            }

            return combinedDataTable;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// 테이블을 로드하는 내부 메서드
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>로드된 DataTable</returns>
        private DataTable LoadInternal(string tableName)
        {
            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                DataTable dataTable = new DataTable(tableName);

                string sql = $"SELECT * FROM [{tableName}]";
                using (var cmd = new SQLiteCommand(sql, Connection))
                using (var adapter = new SQLiteDataAdapter(cmd))
                {
                    adapter.Fill(dataTable);
                }

                return dataTable;
            }
            finally
            {
                if (!wasOpen)
                {
                    CloseConnection();
                }
            }
        }

        #endregion
    }
}
