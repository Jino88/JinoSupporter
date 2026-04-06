using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMaker.R6.SQLProcess
{
    public class BaseSqliteService : IDisposable
    {
        #region Fields & Properties

        private SQLiteConnection _connection;
        private string _dbPath;
        private bool _disposed = false;

        /// <summary>
        /// 데이터베이스 파일 경로
        /// </summary>
        public string DBPath
        {
            get => _dbPath;
            set
            {
                if (_dbPath != value)
                {
                    _dbPath = value;
                    CloseConnection();
                    _connection = null;
                }
            }
        }

        /// <summary>
        /// SQLite 연결 객체
        /// </summary>
        protected SQLiteConnection Connection
        {
            get
            {
                if (_connection == null && !string.IsNullOrEmpty(_dbPath))
                {
                    _connection = new SQLiteConnection($"Data Source={_dbPath};");
                }
                return _connection;
            }
        }

        /// <summary>
        /// 현재 연결이 열려있는지 확인
        /// </summary>
        public bool IsConnectionOpen => Connection?.State == ConnectionState.Open;

        #endregion

        #region Constructors

        /// <summary>
        /// 기본 생성자
        /// </summary>
        protected BaseSqliteService()
        {
        }

        /// <summary>
        /// DB 경로를 지정하는 생성자
        /// </summary>
        /// <param name="dbPath">데이터베이스 파일 경로</param>
        protected BaseSqliteService(string dbPath)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
                throw new ArgumentException("Database path cannot be null or empty.", nameof(dbPath));

            _dbPath = dbPath;
        }

        #endregion

        #region Connection Management

        /// <summary>
        /// 데이터베이스 연결을 엽니다.
        /// </summary>
        public virtual void OpenConnection()
        {
            if (string.IsNullOrEmpty(_dbPath))
                throw new InvalidOperationException("Database path is not set.");

            if (Connection.State == ConnectionState.Closed)
            {
                Connection.Open();
            }
            else if (Connection.State == ConnectionState.Broken)
            {
                Connection.Close();
                Connection.Open();
            }
        }

        /// <summary>
        /// 데이터베이스 연결을 닫습니다.
        /// </summary>
        public virtual void CloseConnection()
        {
            if (Connection?.State == ConnectionState.Open)
            {
                Connection.Close();
            }
        }

        /// <summary>
        /// 연결 상태를 확인하고 필요시 자동으로 엽니다.
        /// </summary>
        protected void EnsureConnectionOpen()
        {
            if (!IsConnectionOpen)
            {
                OpenConnection();
            }
        }

        #endregion

        #region Transaction Support

        /// <summary>
        /// 트랜잭션 내에서 작업을 실행합니다.
        /// </summary>
        /// <param name="action">실행할 작업</param>
        protected void ExecuteInTransaction(Action<SQLiteTransaction> action)
        {
            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                using (var transaction = Connection.BeginTransaction())
                {
                    try
                    {
                        action(transaction);
                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
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
        }

        /// <summary>
        /// 트랜잭션 내에서 작업을 실행하고 결과를 반환합니다.
        /// </summary>
        /// <typeparam name="T">반환 타입</typeparam>
        /// <param name="func">실행할 함수</param>
        /// <returns>함수 실행 결과</returns>
        protected T ExecuteInTransaction<T>(Func<SQLiteTransaction, T> func)
        {
            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                using (var transaction = Connection.BeginTransaction())
                {
                    try
                    {
                        T result = func(transaction);
                        transaction.Commit();
                        return result;
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
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
        }

        #endregion

        #region Command Execution Helpers

        /// <summary>
        /// SQL 명령을 실행합니다 (Non-Query).
        /// </summary>
        /// <param name="sql">실행할 SQL</param>
        /// <param name="parameters">파라미터 딕셔너리</param>
        /// <returns>영향받은 행 수</returns>
        protected int ExecuteNonQuery(string sql, Dictionary<string, object> parameters = null)
        {
            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                using (var cmd = new SQLiteCommand(sql, Connection))
                {
                    AddParameters(cmd, parameters);
                    return cmd.ExecuteNonQuery();
                }
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
        /// SQL 명령을 실행하고 스칼라 값을 반환합니다.
        /// </summary>
        /// <typeparam name="T">반환 타입</typeparam>
        /// <param name="sql">실행할 SQL</param>
        /// <param name="parameters">파라미터 딕셔너리</param>
        /// <returns>쿼리 결과</returns>
        protected T ExecuteScalar<T>(string sql, Dictionary<string, object> parameters = null)
        {
            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                using (var cmd = new SQLiteCommand(sql, Connection))
                {
                    AddParameters(cmd, parameters);
                    object result = cmd.ExecuteScalar();

                    if (result == null || result == DBNull.Value)
                        return default(T);

                    return (T)Convert.ChangeType(result, typeof(T));
                }
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
        /// 명령에 파라미터를 추가합니다.
        /// </summary>
        /// <param name="cmd">SQLite 명령</param>
        /// <param name="parameters">파라미터 딕셔너리</param>
        protected void AddParameters(SQLiteCommand cmd, Dictionary<string, object> parameters)
        {
            if (parameters != null)
            {
                foreach (var param in parameters)
                {
                    cmd.Parameters.AddWithValue(param.Key, param.Value ?? DBNull.Value);
                }
            }
        }

        #endregion

        #region Common Database Operations

        /// <summary>
        /// 테이블이 존재하는지 확인합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>존재 여부</returns>
        public bool TableExists(string tableName)
        {
            string sql = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@tableName;";
            var parameters = new Dictionary<string, object> { { "@tableName", tableName } };

            long count = ExecuteScalar<long>(sql, parameters);
            return count > 0;
        }

        /// <summary>
        /// 테이블의 컬럼 목록을 가져옵니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>컬럼 이름 리스트</returns>
        public List<string> GetTableColumns(string tableName)
        {
            var columns = new List<string>();
            bool wasOpen = IsConnectionOpen;

            try
            {
                EnsureConnectionOpen();

                string sql = $"PRAGMA table_info([{tableName}]);";
                using (var cmd = new SQLiteCommand(sql, Connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(reader["name"].ToString());
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

            return columns;
        }

        /// <summary>
        /// 테이블의 행 개수를 가져옵니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>행 개수</returns>
        public int GetRowCount(string tableName)
        {
            string sql = $"SELECT COUNT(*) FROM [{tableName}];";
            return ExecuteScalar<int>(sql);
        }

        #endregion

        #region IDisposable Implementation

        /// <summary>
        /// 리소스를 해제합니다.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 리소스를 해제합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스 해제 여부</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    CloseConnection();
                    _connection?.Dispose();
                    _connection = null;
                }

                _disposed = true;
            }
        }

        /// <summary>
        /// 소멸자
        /// </summary>
        ~BaseSqliteService()
        {
            Dispose(false);
        }

        #endregion
    }
}
