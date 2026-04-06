using DataMaker.Logger;
using System.Data;
using System.Data.SQLite;

namespace DataMaker.R6.SQLProcess
{
    public interface ISQLiteReader
    {
        string DBPath { get; set; }
        DataTable Read(string tableName);
        DataTable Read(string tableName, List<HashSet<(string ColumnName, string ColumnItem)>> columns);
        List<string> GetColumns(string tableName);
        bool IsTableExist(string tableName);
    }

    /// <summary>
    /// SQLite 데이터베이스에서 데이터를 읽는 클래스
    /// </summary>
    public class clSQLiteReader : BaseSqliteService, ISQLiteReader
    {
        #region Fields

        private readonly ISQLiteLoader _loader;

        #endregion

        #region Constructors

        /// <summary>
        /// Loader를 사용하는 생성자
        /// </summary>
        /// <param name="loader">SQLite 로더</param>
        public clSQLiteReader(ISQLiteLoader loader) : base()
        {
            _loader = loader ?? throw new ArgumentNullException(nameof(loader), "Loader cannot be null.");
            DBPath = loader.DBPath;
        }

        /// <summary>
        /// DB 경로를 직접 지정하는 생성자
        /// </summary>
        /// <param name="dbPath">데이터베이스 경로</param>
        public clSQLiteReader(string dbPath) : base(dbPath)
        {
            _loader = new clSQLiteLoader(dbPath);
        }

        #endregion

        #region ISQLiteReader Implementation

        /// <summary>
        /// 테이블이 존재하는지 확인합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>존재 여부</returns>
        public bool IsTableExist(string tableName)
        {
            return TableExists(tableName);
        }

        /// <summary>
        /// 테이블의 컬럼 목록을 가져옵니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>컬럼 이름 리스트</returns>
        public List<string> GetColumns(string tableName)
        {
            return GetTableColumns(tableName);
        }

        /// <summary>
        /// 테이블 전체를 읽어옵니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>읽어온 DataTable</returns>
        public DataTable Read(string tableName)
        {
            try
            {
                if (_loader != null)
                {
                    return _loader.Load(tableName);
                }
                else
                {
                    // Loader가 없는 경우 직접 읽기
                    return ReadInternal(tableName);
                }
            }
            catch (Exception ex)
            {
                clLogger.Log($"Error reading table '{tableName}': {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 조건에 맞는 데이터를 읽어옵니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="columns">컬럼 조건 목록</param>
        /// <returns>읽어온 DataTable</returns>
        public DataTable Read(string tableName, List<HashSet<(string ColumnName, string ColumnItem)>> columns)
        {
            if (string.IsNullOrEmpty(tableName))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));

            if (columns == null || columns.Count == 0)
                throw new ArgumentException("Columns cannot be null or empty.", nameof(columns));

            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                DataTable dataTable = new DataTable(tableName);

                // WHERE 조건 생성
                var whereConditions = new List<string>();
                var parameters = new Dictionary<string, object>();
                int paramIndex = 0;

                foreach (var columnSet in columns)
                {
                    foreach (var (ColumnName, ColumnItem) in columnSet)
                    {
                        string paramName = $"@param{paramIndex++}";
                        whereConditions.Add($"[{ColumnName}] = {paramName}");
                        parameters.Add(paramName, ColumnItem);
                    }
                }

                string whereClause = whereConditions.Count > 0
                    ? $"WHERE {string.Join(" AND ", whereConditions)}"
                    : "";

                string sql = $"SELECT * FROM [{tableName}] {whereClause}";

                using (var cmd = new SQLiteCommand(sql, Connection))
                {
                    foreach (var param in parameters)
                    {
                        cmd.Parameters.AddWithValue(param.Key, param.Value);
                    }

                    using (var adapter = new SQLiteDataAdapter(cmd))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                return dataTable;
            }
            catch (Exception ex)
            {
                clLogger.Log($"Error reading table '{tableName}' with conditions: {ex.Message}");
                throw;
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

        #region Private Methods

        /// <summary>
        /// 테이블을 직접 읽어오는 내부 메서드
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>읽어온 DataTable</returns>
        private DataTable ReadInternal(string tableName)
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
