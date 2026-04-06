using System.Text.Json;
using System.IO;
using CustomKeyboardCSharp.Models;
using Microsoft.Data.Sqlite;

namespace CustomKeyboardCSharp.Services;

public sealed class TranslationHistoryRepository
{
    private readonly string _dbPath;

    public TranslationHistoryRepository()
    {
        var baseDirectory = CustomKeyboardPathResolver.GetAppDataDirectory();
        Directory.CreateDirectory(baseDirectory);
        _dbPath = Path.Combine(baseDirectory, "translation_history.db");
        Initialize();
    }

    public string DatabasePath => _dbPath;

    public void SaveTranslation(
        AiProvider provider,
        string mode,
        string originalText,
        string directionLabel,
        TranslationBundle result)
    {
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();

        var historyId = InsertHistory(connection, transaction, provider, mode, originalText, directionLabel, result);

        for (var index = 0; index < result.Options.Count; index++)
        {
            var option = result.Options[index];
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText =
                """
                INSERT INTO translation_options (history_id, position, translated_text, nuance)
                VALUES ($historyId, $position, $translatedText, $nuance)
                """;
            command.Parameters.AddWithValue("$historyId", historyId);
            command.Parameters.AddWithValue("$position", index + 1);
            command.Parameters.AddWithValue("$translatedText", option.Text);
            command.Parameters.AddWithValue("$nuance", option.Nuance);
            command.ExecuteNonQuery();
        }

        foreach (var glossary in result.Glossary)
        {
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText =
                """
                INSERT INTO vocabulary
                (created_at, provider, mode, direction, source_word, target_meaning, source_text, history_id)
                VALUES ($createdAt, $provider, $mode, $direction, $sourceWord, $targetMeaning, $sourceText, $historyId)
                """;
            command.Parameters.AddWithValue("$createdAt", DateTimeOffset.UtcNow.ToUnixTimeMilliseconds());
            command.Parameters.AddWithValue("$provider", ProviderStorage(provider));
            command.Parameters.AddWithValue("$mode", mode);
            command.Parameters.AddWithValue("$direction", directionLabel);
            command.Parameters.AddWithValue("$sourceWord", glossary.Word);
            command.Parameters.AddWithValue("$targetMeaning", glossary.Meaning);
            command.Parameters.AddWithValue("$sourceText", originalText);
            command.Parameters.AddWithValue("$historyId", historyId);
            command.ExecuteNonQuery();
        }

        transaction.Commit();
    }

    public List<HistoryItem> GetRecentHistories(int limit = 50)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            SELECT id, created_at, provider, mode, direction, source_text
            FROM translation_history
            ORDER BY created_at DESC
            LIMIT $limit
            """;
        command.Parameters.AddWithValue("$limit", limit);

        using var reader = command.ExecuteReader();
        var items = new List<HistoryItem>();
        while (reader.Read())
        {
            var historyId = reader.GetInt64(0);
            items.Add(new HistoryItem
            {
                Id = historyId,
                CreatedAt = reader.GetInt64(1),
                Provider = reader.GetString(2),
                Mode = reader.GetString(3),
                Direction = reader.GetString(4),
                SourceText = reader.GetString(5),
                Options = GetOptionsForHistory(connection, historyId)
            });
        }

        return items;
    }

    public List<VocabularyItem> GetRecentVocabulary(int limit = 100)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            SELECT id, source_word, target_meaning, direction, provider
            FROM vocabulary
            ORDER BY created_at DESC
            LIMIT $limit
            """;
        command.Parameters.AddWithValue("$limit", limit);

        using var reader = command.ExecuteReader();
        var items = new List<VocabularyItem>();
        while (reader.Read())
        {
            items.Add(new VocabularyItem
            {
                Id = reader.GetInt64(0),
                SourceWord = reader.GetString(1),
                TargetMeaning = reader.GetString(2),
                Direction = reader.GetString(3),
                Provider = reader.GetString(4)
            });
        }

        return items;
    }

    public void DeleteHistory(long historyId)
    {
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();

        ExecuteDelete(connection, transaction, "DELETE FROM translation_options WHERE history_id = $id", historyId);
        ExecuteDelete(connection, transaction, "DELETE FROM vocabulary WHERE history_id = $id", historyId);
        ExecuteDelete(connection, transaction, "DELETE FROM translation_history WHERE id = $id", historyId);

        transaction.Commit();
    }

    public void DeleteVocabulary(long vocabularyId)
    {
        using var connection = OpenConnection();
        ExecuteDelete(connection, null, "DELETE FROM vocabulary WHERE id = $id", vocabularyId);
    }

    public string ExportDatabaseToDownloads()
    {
        var downloads = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Downloads");
        Directory.CreateDirectory(downloads);

        var fileName = $"translation_history_{DateTime.Now:yyyyMMdd_HHmmss}.db";
        var destination = Path.Combine(downloads, fileName);
        using var sourceConnection = OpenConnection();
        var destinationConnectionString = new SqliteConnectionStringBuilder
        {
            DataSource = destination,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared,
            Pooling = true,
            DefaultTimeout = 10
        }.ToString();
        using var destinationConnection = new SqliteConnection(destinationConnectionString);
        destinationConnection.Open();
        sourceConnection.BackupDatabase(destinationConnection);
        return destination;
    }

    public DriveSyncSnapshot ExportSharedSnapshot()
    {
        using var connection = OpenConnection();
        return ExportSharedSnapshot(connection);
    }

    public DriveSyncSnapshot ExportSharedSnapshotFrom(string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Source database file was not found.", sourcePath);
        }

        using var connection = OpenExternalConnection(sourcePath, SqliteOpenMode.ReadOnly, pooling: false);
        ValidateDatabaseSchema(connection);
        return ExportSharedSnapshot(connection);
    }

    public bool TryExportSharedSnapshotFrom(string sourcePath, out DriveSyncSnapshot? snapshot, out string? errorMessage)
    {
        try
        {
            snapshot = ExportSharedSnapshotFrom(sourcePath);
            errorMessage = null;
            return true;
        }
        catch (Exception ex) when (ex is InvalidOperationException or SqliteException or IOException)
        {
            snapshot = null;
            errorMessage = ex.Message;
            return false;
        }
    }

    public void WriteSnapshotToDatabase(string destinationPath, DriveSyncSnapshot snapshot)
    {
        string? directory = Path.GetDirectoryName(destinationPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (File.Exists(destinationPath))
        {
            File.Delete(destinationPath);
        }

        using var connection = OpenExternalConnection(destinationPath, SqliteOpenMode.ReadWriteCreate, pooling: false);
        EnsureSchema(connection);
        ReplaceAllWithSnapshot(connection, snapshot);
    }

    public static DriveSyncSnapshot MergeSnapshots(params DriveSyncSnapshot[] snapshots)
    {
        DriveSyncSnapshot merged = new()
        {
            SchemaVersion = snapshots.Length == 0 ? 3 : snapshots.Max(snapshot => snapshot.SchemaVersion),
            ExportedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
        };

        Dictionary<string, DriveSyncHistoryRecord> historyByKey = new(StringComparer.Ordinal);
        Dictionary<string, long> mergedHistoryIdByKey = new(StringComparer.Ordinal);
        long nextHistoryId = 1;

        foreach (DriveSyncSnapshot snapshot in snapshots)
        {
            foreach (DriveSyncHistoryRecord history in snapshot.Histories.OrderBy(item => item.CreatedAt).ThenBy(item => item.Id))
            {
                string historyKey = BuildHistoryKey(history);
                if (historyByKey.ContainsKey(historyKey))
                {
                    continue;
                }

                DriveSyncHistoryRecord mergedHistory = new()
                {
                    Id = nextHistoryId++,
                    CreatedAt = history.CreatedAt,
                    Provider = history.Provider,
                    Mode = history.Mode,
                    Direction = history.Direction,
                    SourceText = history.SourceText,
                    ResultJson = history.ResultJson
                };

                historyByKey[historyKey] = mergedHistory;
                mergedHistoryIdByKey[historyKey] = mergedHistory.Id;
                merged.Histories.Add(mergedHistory);
            }
        }

        Dictionary<string, DriveSyncOptionRecord> optionByKey = new(StringComparer.Ordinal);
        long nextOptionId = 1;
        foreach (DriveSyncSnapshot snapshot in snapshots)
        {
            Dictionary<long, string> historyKeyBySourceId = snapshot.Histories.ToDictionary(item => item.Id, BuildHistoryKey);
            foreach (DriveSyncOptionRecord option in snapshot.Options.OrderBy(item => item.HistoryId).ThenBy(item => item.Position).ThenBy(item => item.Id))
            {
                if (!historyKeyBySourceId.TryGetValue(option.HistoryId, out string? historyKey) ||
                    !mergedHistoryIdByKey.TryGetValue(historyKey, out long mergedHistoryId))
                {
                    continue;
                }

                string optionKey = $"{historyKey}|{option.Position}|{option.TranslatedText}|{option.Nuance}";
                if (optionByKey.ContainsKey(optionKey))
                {
                    continue;
                }

                DriveSyncOptionRecord mergedOption = new()
                {
                    Id = nextOptionId++,
                    HistoryId = mergedHistoryId,
                    Position = option.Position,
                    TranslatedText = option.TranslatedText,
                    Nuance = option.Nuance
                };

                optionByKey[optionKey] = mergedOption;
                merged.Options.Add(mergedOption);
            }
        }

        Dictionary<string, DriveSyncVocabularyRecord> vocabularyByKey = new(StringComparer.Ordinal);
        long nextVocabularyId = 1;
        foreach (DriveSyncSnapshot snapshot in snapshots)
        {
            Dictionary<long, string> historyKeyBySourceId = snapshot.Histories.ToDictionary(item => item.Id, BuildHistoryKey);
            foreach (DriveSyncVocabularyRecord vocabulary in snapshot.Vocabulary.OrderBy(item => item.CreatedAt).ThenBy(item => item.Id))
            {
                string vocabularyKey = BuildVocabularyKey(vocabulary);
                if (vocabularyByKey.ContainsKey(vocabularyKey))
                {
                    continue;
                }

                long? mergedHistoryId = null;
                if (vocabulary.HistoryId is long sourceHistoryId &&
                    historyKeyBySourceId.TryGetValue(sourceHistoryId, out string? historyKey) &&
                    mergedHistoryIdByKey.TryGetValue(historyKey, out long foundMergedHistoryId))
                {
                    mergedHistoryId = foundMergedHistoryId;
                }

                DriveSyncVocabularyRecord mergedVocabulary = new()
                {
                    Id = nextVocabularyId++,
                    CreatedAt = vocabulary.CreatedAt,
                    Provider = vocabulary.Provider,
                    Mode = vocabulary.Mode,
                    Direction = vocabulary.Direction,
                    SourceWord = vocabulary.SourceWord,
                    TargetMeaning = vocabulary.TargetMeaning,
                    SourceText = vocabulary.SourceText,
                    HistoryId = mergedHistoryId
                };

                vocabularyByKey[vocabularyKey] = mergedVocabulary;
                merged.Vocabulary.Add(mergedVocabulary);
            }
        }

        return merged;
    }

    private static string BuildHistoryKey(DriveSyncHistoryRecord history)
    {
        return $"{history.CreatedAt}|{history.Provider}|{history.Mode}|{history.Direction}|{history.SourceText}|{history.ResultJson}";
    }

    private static string BuildVocabularyKey(DriveSyncVocabularyRecord vocabulary)
    {
        return $"{vocabulary.CreatedAt}|{vocabulary.Provider}|{vocabulary.Mode}|{vocabulary.Direction}|{vocabulary.SourceWord}|{vocabulary.TargetMeaning}|{vocabulary.SourceText}";
    }

    private DriveSyncSnapshot ExportSharedSnapshot(SqliteConnection connection)
    {
        var snapshot = new DriveSyncSnapshot
        {
            SchemaVersion = 3,
            ExportedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
        };

        using (var historyCommand = connection.CreateCommand())
        {
            historyCommand.CommandText =
                """
                SELECT id, created_at, provider, mode, direction, source_text, result_json
                FROM translation_history
                ORDER BY created_at ASC
                """;
            using var reader = historyCommand.ExecuteReader();
            while (reader.Read())
            {
                snapshot.Histories.Add(new DriveSyncHistoryRecord
                {
                    Id = reader.GetInt64(0),
                    CreatedAt = reader.GetInt64(1),
                    Provider = reader.GetString(2),
                    Mode = reader.GetString(3),
                    Direction = reader.GetString(4),
                    SourceText = reader.GetString(5),
                    ResultJson = reader.GetString(6)
                });
            }
        }

        using (var optionCommand = connection.CreateCommand())
        {
            optionCommand.CommandText =
                """
                SELECT id, history_id, position, translated_text, nuance
                FROM translation_options
                ORDER BY history_id ASC, position ASC
                """;
            using var reader = optionCommand.ExecuteReader();
            while (reader.Read())
            {
                snapshot.Options.Add(new DriveSyncOptionRecord
                {
                    Id = reader.GetInt64(0),
                    HistoryId = reader.GetInt64(1),
                    Position = reader.GetInt32(2),
                    TranslatedText = reader.GetString(3),
                    Nuance = reader.GetString(4)
                });
            }
        }

        using (var vocabularyCommand = connection.CreateCommand())
        {
            vocabularyCommand.CommandText =
                """
                SELECT id, created_at, provider, mode, direction, source_word, target_meaning, source_text, history_id
                FROM vocabulary
                ORDER BY created_at ASC
                """;
            using var reader = vocabularyCommand.ExecuteReader();
            while (reader.Read())
            {
                snapshot.Vocabulary.Add(new DriveSyncVocabularyRecord
                {
                    Id = reader.GetInt64(0),
                    CreatedAt = reader.GetInt64(1),
                    Provider = reader.GetString(2),
                    Mode = reader.GetString(3),
                    Direction = reader.GetString(4),
                    SourceWord = reader.GetString(5),
                    TargetMeaning = reader.GetString(6),
                    SourceText = reader.GetString(7),
                    HistoryId = reader.IsDBNull(8) ? null : reader.GetInt64(8)
                });
            }
        }

        return snapshot;
    }

    public void ImportDatabaseFrom(string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Source database file was not found.", sourcePath);
        }

        using var sourceConnection = OpenExternalConnection(sourcePath, SqliteOpenMode.ReadOnly, pooling: false);
        ValidateDatabaseSchema(sourceConnection);
        using var destinationConnection = OpenConnection();
        sourceConnection.BackupDatabase(destinationConnection);
        EnsureSchema(destinationConnection);
    }

    public void ReplaceAllWithSnapshot(DriveSyncSnapshot snapshot)
    {
        using var connection = OpenConnection();
        ReplaceAllWithSnapshot(connection, snapshot);
    }

    private static void ReplaceAllWithSnapshot(SqliteConnection connection, DriveSyncSnapshot snapshot)
    {
        using var transaction = connection.BeginTransaction();

        using (var deleteCommand = connection.CreateCommand())
        {
            deleteCommand.Transaction = transaction;
            deleteCommand.CommandText =
                """
                DELETE FROM translation_options;
                DELETE FROM vocabulary;
                DELETE FROM translation_history;
                """;
            deleteCommand.ExecuteNonQuery();
        }

        foreach (var item in snapshot.Histories)
        {
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText =
                """
                INSERT INTO translation_history (id, created_at, provider, mode, direction, source_text, result_json)
                VALUES ($id, $createdAt, $provider, $mode, $direction, $sourceText, $resultJson)
                """;
            command.Parameters.AddWithValue("$id", item.Id);
            command.Parameters.AddWithValue("$createdAt", item.CreatedAt);
            command.Parameters.AddWithValue("$provider", item.Provider);
            command.Parameters.AddWithValue("$mode", item.Mode);
            command.Parameters.AddWithValue("$direction", item.Direction);
            command.Parameters.AddWithValue("$sourceText", item.SourceText);
            command.Parameters.AddWithValue("$resultJson", item.ResultJson);
            command.ExecuteNonQuery();
        }

        foreach (var item in snapshot.Options)
        {
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText =
                """
                INSERT INTO translation_options (id, history_id, position, translated_text, nuance)
                VALUES ($id, $historyId, $position, $translatedText, $nuance)
                """;
            command.Parameters.AddWithValue("$id", item.Id);
            command.Parameters.AddWithValue("$historyId", item.HistoryId);
            command.Parameters.AddWithValue("$position", item.Position);
            command.Parameters.AddWithValue("$translatedText", item.TranslatedText);
            command.Parameters.AddWithValue("$nuance", item.Nuance);
            command.ExecuteNonQuery();
        }

        foreach (var item in snapshot.Vocabulary)
        {
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText =
                """
                INSERT INTO vocabulary
                (id, created_at, provider, mode, direction, source_word, target_meaning, source_text, history_id)
                VALUES ($id, $createdAt, $provider, $mode, $direction, $sourceWord, $targetMeaning, $sourceText, $historyId)
                """;
            command.Parameters.AddWithValue("$id", item.Id);
            command.Parameters.AddWithValue("$createdAt", item.CreatedAt);
            command.Parameters.AddWithValue("$provider", item.Provider);
            command.Parameters.AddWithValue("$mode", item.Mode);
            command.Parameters.AddWithValue("$direction", item.Direction);
            command.Parameters.AddWithValue("$sourceWord", item.SourceWord);
            command.Parameters.AddWithValue("$targetMeaning", item.TargetMeaning);
            command.Parameters.AddWithValue("$sourceText", item.SourceText);
            command.Parameters.AddWithValue("$historyId", (object?)item.HistoryId ?? DBNull.Value);
            command.ExecuteNonQuery();
        }

        transaction.Commit();
    }

    private void Initialize()
    {
        using var connection = OpenConnection();
        EnsureSchema(connection);
    }

    private long InsertHistory(
        SqliteConnection connection,
        SqliteTransaction transaction,
        AiProvider provider,
        string mode,
        string originalText,
        string directionLabel,
        TranslationBundle result)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText =
            """
            INSERT INTO translation_history (created_at, provider, mode, direction, source_text, result_json)
            VALUES ($createdAt, $provider, $mode, $direction, $sourceText, $resultJson);
            SELECT last_insert_rowid();
            """;
        command.Parameters.AddWithValue("$createdAt", DateTimeOffset.UtcNow.ToUnixTimeMilliseconds());
        command.Parameters.AddWithValue("$provider", ProviderStorage(provider));
        command.Parameters.AddWithValue("$mode", mode);
        command.Parameters.AddWithValue("$direction", directionLabel);
        command.Parameters.AddWithValue("$sourceText", originalText);
        command.Parameters.AddWithValue("$resultJson", JsonSerializer.Serialize(result));
        return (long)(command.ExecuteScalar() ?? 0L);
    }

    private static void ExecuteDelete(SqliteConnection connection, SqliteTransaction? transaction, string sql, long id)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        command.Parameters.AddWithValue("$id", id);
        command.ExecuteNonQuery();
    }

    private static string ProviderStorage(AiProvider provider) => provider switch
    {
        AiProvider.Gemini => "gemini",
        _ => "openai"
    };

    private static List<TranslationOptionItem> GetOptionsForHistory(SqliteConnection connection, long historyId)
    {
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            SELECT position, translated_text, nuance
            FROM translation_options
            WHERE history_id = $historyId
            ORDER BY position ASC
            """;
        command.Parameters.AddWithValue("$historyId", historyId);

        using var reader = command.ExecuteReader();
        var options = new List<TranslationOptionItem>();
        while (reader.Read())
        {
            options.Add(new TranslationOptionItem
            {
                Position = reader.GetInt32(0),
                TranslatedText = reader.GetString(1),
                Nuance = reader.GetString(2)
            });
        }

        return options;
    }

    private SqliteConnection OpenConnection()
    {
        return OpenExternalConnection(_dbPath, SqliteOpenMode.ReadWriteCreate, pooling: true);
    }

    private static SqliteConnection OpenExternalConnection(string dbPath, SqliteOpenMode mode, bool pooling)
    {
        var connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = dbPath,
            Mode = mode,
            Cache = SqliteCacheMode.Shared,
            Pooling = pooling,
            DefaultTimeout = 10
        }.ToString();

        var connection = new SqliteConnection(connectionString);
        connection.Open();

        using var pragmaCommand = connection.CreateCommand();
        pragmaCommand.CommandText = "PRAGMA busy_timeout=5000;";
        pragmaCommand.ExecuteNonQuery();

        return connection;
    }

    private static void ValidateDatabaseSchema(SqliteConnection connection)
    {
        string[] requiredTables =
        [
            "translation_history",
            "translation_options",
            "vocabulary"
        ];

        foreach (string tableName in requiredTables)
        {
            using var command = connection.CreateCommand();
            command.CommandText =
                """
                SELECT COUNT(*)
                FROM sqlite_master
                WHERE type = 'table' AND name = $tableName
                """;
            command.Parameters.AddWithValue("$tableName", tableName);

            long count = (long)(command.ExecuteScalar() ?? 0L);
            if (count == 0)
            {
                throw new InvalidOperationException(
                    $"Downloaded DB is invalid. Required table '{tableName}' was not found.");
            }
        }
    }

    private static void EnsureSchema(SqliteConnection connection)
    {
        using (var pragmaCommand = connection.CreateCommand())
        {
            pragmaCommand.CommandText =
                """
                PRAGMA journal_mode=WAL;
                PRAGMA synchronous=NORMAL;
                PRAGMA busy_timeout=5000;
                """;
            pragmaCommand.ExecuteNonQuery();
        }

        using var command = connection.CreateCommand();
        command.CommandText =
            """
            CREATE TABLE IF NOT EXISTS translation_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at INTEGER NOT NULL,
                provider TEXT NOT NULL,
                mode TEXT NOT NULL,
                direction TEXT NOT NULL,
                source_text TEXT NOT NULL,
                result_json TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS vocabulary (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at INTEGER NOT NULL,
                provider TEXT NOT NULL,
                mode TEXT NOT NULL,
                direction TEXT NOT NULL,
                source_word TEXT NOT NULL,
                target_meaning TEXT NOT NULL,
                source_text TEXT NOT NULL,
                history_id INTEGER
            );

            CREATE TABLE IF NOT EXISTS translation_options (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                history_id INTEGER NOT NULL,
                position INTEGER NOT NULL,
                translated_text TEXT NOT NULL,
                nuance TEXT NOT NULL
            );
            """;
        command.ExecuteNonQuery();
    }
}
