using DiskTree.Models;
using Microsoft.Data.Sqlite;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace DiskTree.Services;

public readonly record struct IndexWriteProgress(int WrittenCount, int TotalCount);

public sealed class DiskIndexStore : IDisposable
{
    private readonly SqliteConnection _connection;

    public DiskIndexStore(string databasePath)
    {
        string normalizedPath = Path.GetFullPath(databasePath);
        string? directory = Path.GetDirectoryName(normalizedPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        _connection = new SqliteConnection($"Data Source={normalizedPath}");
        _connection.Open();
        ConfigureConnectionForFastBulkWrite();
        EnsureSchema();
    }

    public void UpsertFiles(
        IEnumerable<IndexedFileRecord> files,
        CancellationToken cancellationToken,
        int totalCount = 0,
        IProgress<IndexWriteProgress>? progress = null)
    {
        using var transaction = _connection.BeginTransaction();
        using var command = _connection.CreateCommand();

        command.Transaction = transaction;
        command.CommandText =
            """
            INSERT INTO FileIndex
            (
                FilePath,
                FileName,
                DirectoryPath,
                FileSize,
                HeadTailHash,
                LastWriteUtc,
                ScannedUtc
            )
            VALUES
            (
                @filePath,
                @fileName,
                @directoryPath,
                @fileSize,
                @headTailHash,
                @lastWriteUtc,
                @scannedUtc
            )
            ON CONFLICT(FilePath) DO UPDATE SET
                FileName = excluded.FileName,
                DirectoryPath = excluded.DirectoryPath,
                FileSize = excluded.FileSize,
                HeadTailHash = excluded.HeadTailHash,
                LastWriteUtc = excluded.LastWriteUtc,
                ScannedUtc = excluded.ScannedUtc;
            """;

        SqliteParameter pathParam = command.Parameters.Add("@filePath", SqliteType.Text);
        SqliteParameter nameParam = command.Parameters.Add("@fileName", SqliteType.Text);
        SqliteParameter directoryParam = command.Parameters.Add("@directoryPath", SqliteType.Text);
        SqliteParameter sizeParam = command.Parameters.Add("@fileSize", SqliteType.Integer);
        SqliteParameter hashParam = command.Parameters.Add("@headTailHash", SqliteType.Text);
        SqliteParameter lastWriteParam = command.Parameters.Add("@lastWriteUtc", SqliteType.Text);
        SqliteParameter scannedParam = command.Parameters.Add("@scannedUtc", SqliteType.Text);
        command.Prepare();

        int writtenCount = 0;

        foreach (IndexedFileRecord file in files)
        {
            cancellationToken.ThrowIfCancellationRequested();

            pathParam.Value = file.FilePath;
            nameParam.Value = file.FileName;
            directoryParam.Value = file.DirectoryPath;
            sizeParam.Value = file.FileSize;
            hashParam.Value = file.HeadTailHash;
            lastWriteParam.Value = file.LastWriteUtc.ToString("O", CultureInfo.InvariantCulture);
            scannedParam.Value = file.ScannedUtc.ToString("O", CultureInfo.InvariantCulture);

            command.ExecuteNonQuery();
            writtenCount++;

            if (progress is not null && (writtenCount % 250 == 0 || (totalCount > 0 && writtenCount == totalCount)))
            {
                progress.Report(new IndexWriteProgress(writtenCount, totalCount));
            }
        }

        transaction.Commit();
        progress?.Report(new IndexWriteProgress(writtenCount, totalCount));
    }

    public IReadOnlyList<IndexedFileRecord> FindCandidates(long fileSize, string headTailHash, string excludeFilePath)
    {
        using var command = _connection.CreateCommand();
        command.CommandText =
            """
            SELECT
                FilePath,
                FileName,
                DirectoryPath,
                FileSize,
                HeadTailHash,
                LastWriteUtc,
                ScannedUtc
            FROM FileIndex
            WHERE FileSize = @fileSize
              AND HeadTailHash = @headTailHash
              AND FilePath <> @excludeFilePath
            ORDER BY FilePath;
            """;

        command.Parameters.AddWithValue("@fileSize", fileSize);
        command.Parameters.AddWithValue("@headTailHash", headTailHash);
        command.Parameters.AddWithValue("@excludeFilePath", Path.GetFullPath(excludeFilePath));

        var results = new List<IndexedFileRecord>();
        using SqliteDataReader reader = command.ExecuteReader();
        while (reader.Read())
        {
            string filePath = reader.GetString(0);
            string fileName = reader.GetString(1);
            string directoryPath = reader.GetString(2);
            long size = reader.GetInt64(3);
            string hash = reader.GetString(4);

            DateTime lastWriteUtc = DateTime.Parse(reader.GetString(5), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
            DateTime scannedUtc = DateTime.Parse(reader.GetString(6), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);

            results.Add(new IndexedFileRecord(
                filePath,
                fileName,
                directoryPath,
                size,
                hash,
                lastWriteUtc,
                scannedUtc));
        }

        return results;
    }

    public int GetIndexedFileCount()
    {
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM FileIndex;";
        object? value = command.ExecuteScalar();
        if (value is long longCount)
        {
            return checked((int)longCount);
        }

        if (value is int intCount)
        {
            return intCount;
        }

        return 0;
    }

    public IReadOnlyDictionary<string, IndexedFileSnapshot> GetIndexedMetadataByRoot(string rootPath)
    {
        string normalizedRootPath = NormalizeRootPath(rootPath);
        string prefixPath = BuildPrefixPath(normalizedRootPath);

        using var command = _connection.CreateCommand();
        command.CommandText =
            """
            SELECT
                FilePath,
                FileSize,
                LastWriteUtc,
                HeadTailHash
            FROM FileIndex
            WHERE FilePath = @exactPath
               OR instr(FilePath, @prefixPath) = 1;
            """;

        command.Parameters.AddWithValue("@exactPath", normalizedRootPath);
        command.Parameters.AddWithValue("@prefixPath", prefixPath);

        var results = new Dictionary<string, IndexedFileSnapshot>(StringComparer.OrdinalIgnoreCase);
        using SqliteDataReader reader = command.ExecuteReader();
        while (reader.Read())
        {
            string filePath = reader.GetString(0);
            long fileSize = reader.GetInt64(1);
            DateTime lastWriteUtc = DateTime.Parse(reader.GetString(2), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
            string headTailHash = reader.GetString(3);
            results[filePath] = new IndexedFileSnapshot(fileSize, lastWriteUtc, headTailHash);
        }

        return results;
    }

    public int DeleteMissingPathsUnderRoot(string rootPath, IReadOnlySet<string> existingPaths, CancellationToken cancellationToken)
    {
        string normalizedRootPath = NormalizeRootPath(rootPath);
        string prefixPath = BuildPrefixPath(normalizedRootPath);

        var indexedPaths = new List<string>(capacity: 4096);

        using (SqliteCommand readCommand = _connection.CreateCommand())
        {
            readCommand.CommandText =
                """
                SELECT FilePath
                FROM FileIndex
                WHERE FilePath = @exactPath
                   OR instr(FilePath, @prefixPath) = 1;
                """;

            readCommand.Parameters.AddWithValue("@exactPath", normalizedRootPath);
            readCommand.Parameters.AddWithValue("@prefixPath", prefixPath);

            using SqliteDataReader reader = readCommand.ExecuteReader();
            while (reader.Read())
            {
                indexedPaths.Add(reader.GetString(0));
            }
        }

        if (indexedPaths.Count == 0)
        {
            return 0;
        }

        using var transaction = _connection.BeginTransaction();
        using var deleteCommand = _connection.CreateCommand();
        deleteCommand.Transaction = transaction;
        deleteCommand.CommandText = "DELETE FROM FileIndex WHERE FilePath = @filePath;";
        SqliteParameter pathParam = deleteCommand.Parameters.Add("@filePath", SqliteType.Text);
        deleteCommand.Prepare();

        int deletedCount = 0;
        foreach (string indexedPath in indexedPaths)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (existingPaths.Contains(indexedPath))
            {
                continue;
            }

            pathParam.Value = indexedPath;
            deletedCount += deleteCommand.ExecuteNonQuery();
        }

        transaction.Commit();
        return deletedCount;
    }

    public void Dispose()
    {
        _connection.Dispose();
    }

    private void EnsureSchema()
    {
        using var command = _connection.CreateCommand();
        command.CommandText =
            """
            CREATE TABLE IF NOT EXISTS FileIndex
            (
                FilePath TEXT PRIMARY KEY,
                FileName TEXT NOT NULL,
                DirectoryPath TEXT NOT NULL,
                FileSize INTEGER NOT NULL,
                HeadTailHash TEXT NOT NULL,
                LastWriteUtc TEXT NOT NULL,
                ScannedUtc TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS IX_FileIndex_SizeHash
                ON FileIndex(FileSize, HeadTailHash);

            CREATE INDEX IF NOT EXISTS IX_FileIndex_FileName
                ON FileIndex(FileName);
            """;
        command.ExecuteNonQuery();
    }

    private void ConfigureConnectionForFastBulkWrite()
    {
        using var command = _connection.CreateCommand();
        command.CommandText =
            """
            PRAGMA journal_mode = WAL;
            PRAGMA synchronous = NORMAL;
            PRAGMA temp_store = MEMORY;
            PRAGMA cache_size = -200000;
            """;
        command.ExecuteNonQuery();
    }

    private static string NormalizeRootPath(string rootPath)
    {
        string normalized = Path.GetFullPath(rootPath);
        string root = Path.GetPathRoot(normalized) ?? string.Empty;
        if (normalized.Length > root.Length)
        {
            normalized = normalized.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        return normalized;
    }

    private static string BuildPrefixPath(string normalizedRootPath)
    {
        return normalizedRootPath.EndsWith(Path.DirectorySeparatorChar)
            ? normalizedRootPath
            : normalizedRootPath + Path.DirectorySeparatorChar;
    }
}
