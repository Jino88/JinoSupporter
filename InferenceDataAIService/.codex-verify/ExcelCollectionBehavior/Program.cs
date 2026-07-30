using InferenceDataAIService.Wpf;
using Microsoft.Data.Sqlite;

var verificationRoot = Path.GetFullPath(
    Path.Combine(
        AppContext.BaseDirectory,
        "runtime-" + Guid.NewGuid().ToString("N")));
var sourceRoot = Path.Combine(verificationRoot, "source");
var secondSourceRoot = Path.Combine(
    verificationRoot,
    "second-source");
var archiveRoot = Path.Combine(verificationRoot, "archive");

try
{
    Directory.CreateDirectory(Path.Combine(sourceRoot, "a"));
    Directory.CreateDirectory(Path.Combine(sourceRoot, "b"));
    Directory.CreateDirectory(secondSourceRoot);
    Directory.CreateDirectory(archiveRoot);

    var firstSource = Path.Combine(sourceRoot, "a", "same.xlsx");
    var secondSource = Path.Combine(sourceRoot, "b", "same.xlsx");
    var thirdSource = Path.Combine(sourceRoot, "b", "other.xlsx");
    var searchedExisting = Path.Combine(
        sourceRoot,
        "linked.xlsx");
    var secondRootFile = Path.Combine(
        secondSourceRoot,
        "second-root.xlsx");
    var databaseLinkedSource = Path.Combine(
        verificationRoot,
        "linked_1234567890_clean.xlsx");
    await File.WriteAllTextAsync(firstSource, "first");
    await File.WriteAllTextAsync(secondSource, "second");
    await File.WriteAllTextAsync(thirdSource, "third");
    await File.WriteAllTextAsync(searchedExisting, "search");
    await File.WriteAllTextAsync(secondRootFile, "second root");
    await File.WriteAllTextAsync(databaseLinkedSource, "database");

    var databasePath = Path.Combine(
        verificationRoot,
        "search.sqlite");
    using (var connection = new SqliteConnection(
               $"Data Source={databasePath}"))
    {
        await connection.OpenAsync();
        using var command = connection.CreateCommand();
        command.CommandText = """
            CREATE TABLE source_documents(
                document_id INTEGER PRIMARY KEY,
                source_path TEXT NOT NULL,
                original_file_name TEXT,
                lifecycle_status TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE source_revisions(
                document_id INTEGER NOT NULL,
                is_current INTEGER NOT NULL
            );
            INSERT INTO source_documents(
                document_id,
                source_path,
                original_file_name,
                lifecycle_status,
                updated_at
            ) VALUES (
                1,
                $sourcePath,
                'linked.xlsx',
                'ACTIVE',
                '2026-07-25T00:00:00Z'
            );
            INSERT INTO source_revisions(document_id, is_current)
            VALUES (1, 1);
            """;
        command.Parameters.AddWithValue(
            "$sourcePath",
            databaseLinkedSource);
        await command.ExecuteNonQueryAsync();
    }

    var searchResult = await ExcelFolderSearchService.SearchAsync(
        sourceRoot,
        databasePath);
    var existingRow = searchResult.Rows.Single(
        row => row.FileName == "linked.xlsx");
    Assert(
        existingRow.ExistsInDatabase,
        "The matching search file must be marked DB existing.");
    Assert(
        string.Equals(
            existingRow.DatabaseSourcePath,
            Path.GetFullPath(databaseLinkedSource),
            StringComparison.OrdinalIgnoreCase),
        "The DB-existing row must expose the linked DB source path.");

    var multiSearchResult =
        await ExcelFolderSearchService.SearchManyAsync(
            [sourceRoot, secondSourceRoot, sourceRoot],
            databasePath);
    Assert(
        multiSearchResult.Rows.Count == searchResult.Rows.Count + 1,
        "Multiple roots must be combined and duplicate roots ignored.");
    Assert(
        multiSearchResult.Rows.Any(row =>
            string.Equals(
                row.SearchRoot,
                Path.GetFullPath(secondSourceRoot),
                StringComparison.OrdinalIgnoreCase)
            && row.FileName == "second-root.xlsx"),
        "Each result must expose the root where it was discovered.");

    var rows = new[]
    {
        Row(firstSource, "a"),
        Row(secondSource, "b"),
        Row(thirdSource, "b"),
    };

    var firstResult = await ExcelLocalCopyService.CopyAsync(
        rows,
        archiveRoot);
    Assert(
        string.Equals(
            Path.GetFullPath(archiveRoot),
            firstResult.LocalRoot,
            StringComparison.OrdinalIgnoreCase),
        "Result root must be the configured archive itself.");
    Assert(
        Directory.GetDirectories(archiveRoot).Length == 0,
        "The archive must not contain generated subdirectories.");
    Assert(
        Directory.GetFiles(archiveRoot).Length == 3,
        "All three source files must be stored directly in the archive.");
    Assert(
        File.Exists(Path.Combine(archiveRoot, "same.xlsx")),
        "The first same-name file is missing.");
    Assert(
        File.Exists(Path.Combine(archiveRoot, "same (2).xlsx")),
        "A collision-safe flat filename was not created.");

    await ExcelLocalCopyService.CopyAsync(rows, archiveRoot);
    Assert(
        Directory.GetFiles(archiveRoot).Length == 3,
        "Repeating the same collection must not duplicate identical files.");
    Assert(
        Directory.GetDirectories(archiveRoot).Length == 0,
        "A repeated collection must not create subdirectories.");

    Console.WriteLine(
        "PASS: DB status links to source_path; configured archive root "
        + "only; multiple search roots; no subfolders; flat collision "
        + "handling and repeat deduplication verified.");
}
finally
{
    SqliteConnection.ClearAllPools();
    if (Directory.Exists(verificationRoot))
        Directory.Delete(verificationRoot, recursive: true);
}

static ExcelFolderSearchRow Row(
    string sourcePath,
    string relativeFolder) =>
    new(
        "DB 없음",
        Path.GetFileName(sourcePath),
        Path.GetDirectoryName(sourcePath) ?? string.Empty,
        relativeFolder,
        sourcePath,
        Path.Combine(relativeFolder, Path.GetFileName(sourcePath)),
        false,
        string.Empty);

static void Assert(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}
