using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// The single SQLite file the app owns: <c>AppStoragePaths.Combine("PortalGNLAB.db")</c>.
/// Every service that used to keep its own .db file opens its connection here.
/// Per-request scratch files (NG Rate temp_*.db, fcost_yyyyMMdd_HHmmss.db) are not stores
/// and stay outside; <see cref="PortalDbMigrator"/> moves the old stores in once.
/// </summary>
public static class PortalDb
{
    public const string FileName = "PortalGNLAB.db";

    /// <summary>Default location; <c>PORTALGNLAB_DATA_DIR</c> relocates it with the rest of the data folder.</summary>
    public static string FilePath => AppStoragePaths.Combine(FileName);

    public static string ConnectionString(bool readOnly = false) => ConnectionString(FilePath, readOnly);

    public static string ConnectionString(string path, bool readOnly = false) =>
        new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
            DefaultTimeout = 30,
        }.ToString();

    public static SqliteConnection Open(bool readOnly = false) => Open(FilePath, readOnly);

    public static SqliteConnection Open(string path, bool readOnly = false)
    {
        if (!readOnly)
        {
            string? dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
        }

        var conn = new SqliteConnection(ConnectionString(path, readOnly));
        conn.Open();
        return conn;
    }
}
