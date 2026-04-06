using System.Data;

namespace DataMaker.R6.SQLProcess
{
    /// <summary>
    /// SQLite 파일 입출력을 위한 통합 인터페이스 클래스
    /// Loader, Reader, Writer, Processor를 모두 통합하여 사용하기 쉽게 제공합니다.
    /// </summary>
    public class clSQLFileIO : IDisposable
    {
        #region Fields & Properties

        private bool _disposed = false;

        /// <summary>
        /// SQLite Loader
        /// </summary>
        public clSQLiteLoader Loader { get; private set; }

        /// <summary>
        /// SQLite Writer
        /// </summary>
        public clSQLiteWriter Writer { get; private set; }

        /// <summary>
        /// SQLite Data Processor
        /// </summary>
        public clSQLiteDataProcessing Processor { get; private set; }

        /// <summary>
        /// SQLite Reader
        /// </summary>
        public clSQLiteReader Reader { get; private set; }

        /// <summary>
        /// 데이터베이스 경로
        /// </summary>
        public string DBPath
        {
            get => Loader?.DBPath;
            set
            {
                if (Loader != null) Loader.DBPath = value;
                if (Writer != null) Writer.DBPath = value;
                if (Processor != null) Processor.DBPath = value;
                if (Reader != null) Reader.DBPath = value;
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// DB 경로를 지정하는 생성자
        /// </summary>
        /// <param name="dbPath">데이터베이스 파일 경로</param>
        public clSQLFileIO(string dbPath)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
                throw new ArgumentException("Database path cannot be null or empty.", nameof(dbPath));

            Initialize(dbPath);
        }

        #endregion

        #region Initialization

        /// <summary>
        /// 모든 컴포넌트를 초기화합니다.
        /// </summary>
        private void Initialize(string dbPath)
        {
            Loader = new clSQLiteLoader(dbPath);
            Writer = new clSQLiteWriter(dbPath);
            Reader = new clSQLiteReader(Loader);
            Processor = new clSQLiteDataProcessing(Loader, Reader, Writer);
        }

        #endregion

        #region Load Methods

        /// <summary>
        /// 테이블을 로드합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>로드된 DataTable</returns>
        public DataTable LoadTable(string tableName)
        {
            return Loader.Load(tableName);
        }

        /// <summary>
        /// 지정된 경로의 데이터베이스에서 테이블을 로드합니다.
        /// </summary>
        /// <param name="dbPath">데이터베이스 경로</param>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>로드된 DataTable</returns>
        public DataTable LoadTable(string dbPath, string tableName)
        {
            string originalPath = DBPath;
            try
            {
                DBPath = dbPath;
                return Loader.Load(tableName);
            }
            finally
            {
                DBPath = originalPath;
            }
        }

        /// <summary>
        /// 조건에 맞는 데이터를 로드합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="modelName">모델 이름 (조건 값)</param>
        /// <param name="columnName">컬럼 이름 (조건 컬럼)</param>
        /// <returns>로드된 DataTable</returns>
        public DataTable LoadTable(string tableName, string modelName, string columnName)
        {
            return Loader.Load(tableName, modelName, columnName);
        }

        /// <summary>
        /// 여러 테이블을 로드하여 병합합니다.
        /// </summary>
        /// <param name="tableNames">테이블 이름 목록</param>
        /// <returns>병합된 DataTable</returns>
        public DataTable LoadTables(List<string> tableNames)
        {
            return Loader.Load(tableNames);
        }

        #endregion

        #region Write Methods

        /// <summary>
        /// 테이블을 생성하고 데이터를 씁니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="table">쓸 데이터</param>
        public void WriteTable(string tableName, DataTable table)
        {
            Writer.Write(tableName, table);
        }

        /// <summary>
        /// 지정된 경로의 데이터베이스에 테이블을 씁니다.
        /// </summary>
        /// <param name="dbPath">데이터베이스 경로</param>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="table">쓸 데이터</param>
        public void WriteTable(string dbPath, string tableName, DataTable table)
        {
            string originalPath = DBPath;
            try
            {
                DBPath = dbPath;
                Writer.Write(tableName, table);
            }
            finally
            {
                DBPath = originalPath;
            }
        }

        /// <summary>
        /// 기존 테이블에 데이터를 추가합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="table">추가할 데이터</param>
        public void AppendTable(string tableName, DataTable table)
        {
            Writer.Append(tableName, table);
        }

        #endregion

        #region Read Methods

        /// <summary>
        /// 테이블을 읽어옵니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>읽어온 DataTable</returns>
        public DataTable ReadTable(string tableName)
        {
            return Reader.Read(tableName);
        }

        /// <summary>
        /// 테이블의 컬럼 목록을 가져옵니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>컬럼 이름 리스트</returns>
        public List<string> GetColumns(string tableName)
        {
            return Reader.GetColumns(tableName);
        }

        /// <summary>
        /// 테이블이 존재하는지 확인합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>존재 여부</returns>
        public bool IsTableExist(string tableName)
        {
            return Reader.IsTableExist(tableName);
        }

        #endregion

        #region Processing Methods

        /// <summary>
        /// 테이블을 생성합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="columnNames">컬럼 이름 배열</param>
        /// <param name="columnTypes">컬럼 타입 배열</param>
        public void CreateTable(string tableName, string[] columnNames, string[] columnTypes)
        {
            Processor.CreateTable(tableName, columnNames, columnTypes);
        }

        /// <summary>
        /// 테이블을 삭제합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        public void DropTable(string tableName)
        {
            Processor.DropTable(tableName);
        }

        /// <summary>
        /// 두 테이블의 공통 컬럼 데이터를 복사합니다.
        /// </summary>
        /// <param name="sourceTable">원본 테이블</param>
        /// <param name="targetTable">대상 테이블</param>
        public void CopyCommonColumns(string sourceTable, string targetTable)
        {
            Processor.CopyCommonColumns(sourceTable, targetTable);
        }

        /// <summary>
        /// 컬럼의 고유 값 목록을 반환합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="columnName">컬럼 이름</param>
        /// <returns>고유 값 리스트</returns>
        public List<string> GetUniqueValues(string tableName, string columnName)
        {
            return Processor.GetListUniqueValuesOfColumns(tableName, columnName);
        }

        /// <summary>
        /// 테이블의 행 개수를 반환합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <returns>행 개수</returns>
        public int GetRowCount(string tableName)
        {
            return Processor.GetRowCount(tableName);
        }

        /// <summary>
        /// 다른 테이블의 데이터로 현재 테이블을 업데이트합니다.
        /// </summary>
        /// <param name="targetTable">대상 테이블</param>
        /// <param name="sourceTable">원본 테이블</param>
        /// <param name="matchColumns">매칭 컬럼</param>
        /// <param name="updateColumns">업데이트할 컬럼</param>
        public void UpdateTableFromTable(
            string targetTable,
            string sourceTable,
            (string TargetCol, string SourceCol)[] matchColumns,
            (string TargetCol, string SourceCol)[] updateColumns)
        {
            Processor.UpdateTableFromTable(targetTable, sourceTable, matchColumns, updateColumns);
        }

        /// <summary>
        /// 데이터를 삽입합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="data">삽입할 데이터</param>
        public void InsertData(string tableName, Dictionary<string, object> data)
        {
            Processor.InsertData(tableName, data);
        }

        /// <summary>
        /// 조건에 맞는 데이터를 업데이트합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="data">업데이트할 데이터</param>
        /// <param name="condition">WHERE 조건</param>
        public void UpdateData(string tableName, Dictionary<string, object> data, string condition)
        {
            Processor.UpdateData(tableName, data, condition);
        }

        /// <summary>
        /// 조건에 맞는 데이터를 삭제합니다.
        /// </summary>
        /// <param name="tableName">테이블 이름</param>
        /// <param name="condition">WHERE 조건</param>
        public void DeleteData(string tableName, string condition)
        {
            Processor.DeleteData(tableName, condition);
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
                    Loader?.Dispose();
                    Writer?.Dispose();
                    Processor?.Dispose();
                    Reader?.Dispose();

                    Loader = null;
                    Writer = null;
                    Processor = null;
                    Reader = null;
                }

                _disposed = true;
            }
        }

        /// <summary>
        /// 소멸자
        /// </summary>
        ~clSQLFileIO()
        {
            Dispose(false);
        }

        #endregion
    }
}
