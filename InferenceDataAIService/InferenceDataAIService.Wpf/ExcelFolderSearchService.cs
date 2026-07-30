using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

internal static class ExcelFolderSearchService
{
    private static readonly HashSet<string> ExcelExtensions = new(
        [".xlsx", ".xlsm", ".xlsb", ".xls"],
        StringComparer.OrdinalIgnoreCase);

    private static readonly EnumerationOptions ReadOnlyEnumerationOptions =
        new()
        {
            RecurseSubdirectories = true,
            IgnoreInaccessible = true,
            ReturnSpecialDirectories = false,
            AttributesToSkip = FileAttributes.ReparsePoint,
        };

    internal static Task<ExcelFolderSearchResult> SearchAsync(
        string rootPath,
        string databasePath) =>
        SearchManyAsync([rootPath], databasePath);

    internal static Task<ExcelFolderSearchResult> SearchManyAsync(
        IReadOnlyList<string> rootPaths,
        string databasePath) =>
        Task.Run(() => SearchMany(rootPaths, databasePath));

    internal static string NormalizeWorkbookFileName(string value)
    {
        var fileName = Path.GetFileName(value);
        if (string.IsNullOrWhiteSpace(fileName))
            return string.Empty;

        var extension = Path.GetExtension(fileName);
        var stem = Path.GetFileNameWithoutExtension(fileName);
        stem = Regex.Replace(
            stem,
            "_clean$",
            string.Empty,
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        stem = Regex.Replace(
            stem,
            @"(?:_[0-9]{9,13})+$",
            string.Empty,
            RegexOptions.CultureInvariant);
        return (stem.Trim() + extension).ToLowerInvariant();
    }

    private static ExcelFolderSearchResult SearchMany(
        IReadOnlyList<string> rootPaths,
        string databasePath)
    {
        var normalizedRoots = rootPaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(path => Path.GetFullPath(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (normalizedRoots.Length == 0)
            throw new ArgumentException(
                "검색할 폴더를 하나 이상 지정하세요.",
                nameof(rootPaths));
        foreach (var root in normalizedRoots)
        {
            if (!Directory.Exists(root))
                throw new DirectoryNotFoundException(
                    $"검색 폴더를 찾을 수 없습니다: {root}");
        }

        var databaseFiles = ReadDatabaseFiles(databasePath);
        var rows = new List<ExcelFolderSearchRow>();
        var seenFiles = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var normalizedRoot in normalizedRoots)
        {
            foreach (var discoveredPath in Directory.EnumerateFiles(
                         normalizedRoot,
                         "*",
                         ReadOnlyEnumerationOptions))
            {
                if (!ExcelExtensions.Contains(
                        Path.GetExtension(discoveredPath))
                    || Path.GetFileName(discoveredPath).StartsWith(
                        "~$",
                        StringComparison.Ordinal))
                {
                    continue;
                }

                var path = Path.GetFullPath(discoveredPath);
                if (!seenFiles.Add(path))
                    continue;
                var fileName = Path.GetFileName(path);
                var parent = Path.GetDirectoryName(path)
                    ?? normalizedRoot;
                var relativeFolder = Path.GetRelativePath(
                    normalizedRoot,
                    parent);
                if (string.Equals(
                        relativeFolder,
                        ".",
                        StringComparison.Ordinal))
                {
                    relativeFolder = "선택 폴더";
                }
                var normalizedFileName =
                    NormalizeWorkbookFileName(fileName);
                var exists = databaseFiles.TryGetValue(
                    normalizedFileName,
                    out var databaseSourcePath);
                rows.Add(new ExcelFolderSearchRow(
                    exists ? "DB 있음" : "DB 없음",
                    fileName,
                    normalizedRoot,
                    relativeFolder,
                    path,
                    Path.GetRelativePath(normalizedRoot, path),
                    exists,
                    databaseSourcePath ?? string.Empty));
            }
        }

        var orderedRows = rows
            .OrderBy(row => row.ExistsInDatabase)
            .ThenBy(
                row => row.FileName,
                StringComparer.OrdinalIgnoreCase)
            .ThenBy(
                row => row.SearchRoot,
                StringComparer.OrdinalIgnoreCase)
            .ThenBy(
                row => row.RelativeFolder,
                StringComparer.OrdinalIgnoreCase)
            .ToArray();

        return new ExcelFolderSearchResult(
            normalizedRoots.Length == 1
                ? normalizedRoots[0]
                : $"{normalizedRoots.Length:N0}개 검색 폴더",
            orderedRows,
            orderedRows.Count(row => row.ExistsInDatabase),
            orderedRows.Count(row => !row.ExistsInDatabase));
    }

    private static Dictionary<string, string> ReadDatabaseFiles(
        string databasePath)
    {
        if (!File.Exists(databasePath))
            throw new FileNotFoundException(
                "현재 적재 DB를 찾을 수 없습니다.",
                databasePath);

        using var connection = new SqliteConnection(
            new SqliteConnectionStringBuilder
            {
                DataSource = Path.GetFullPath(databasePath),
                Mode = SqliteOpenMode.ReadOnly,
                Cache = SqliteCacheMode.Private,
            }.ToString());
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                document.source_path,
                document.original_file_name
            FROM source_documents document
            JOIN source_revisions revision
              ON revision.document_id=document.document_id
             AND revision.is_current=1
            WHERE document.lifecycle_status='ACTIVE'
            ORDER BY document.updated_at DESC, document.document_id DESC;
            """;
        var result = new Dictionary<string, string>(
            StringComparer.Ordinal);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var sourcePath = reader.GetString(0);
            AddDatabaseFile(result, sourcePath, sourcePath);
            if (!reader.IsDBNull(1))
                AddDatabaseFile(
                    result,
                    reader.GetString(1),
                    sourcePath);
        }
        return result;
    }

    private static void AddDatabaseFile(
        IDictionary<string, string> destination,
        string matchValue,
        string sourcePath)
    {
        var normalized = NormalizeWorkbookFileName(matchValue);
        if (normalized.Length > 0
            && !destination.ContainsKey(normalized))
        {
            destination[normalized] =
                Path.GetFullPath(sourcePath);
        }
    }
}

internal sealed record ExcelFolderSearchResult(
    string RootPath,
    IReadOnlyList<ExcelFolderSearchRow> Rows,
    int ExistingCount,
    int MissingCount);

internal sealed record ExcelFolderSearchRow(
    string DatabaseStatus,
    string FileName,
    string SearchRoot,
    string RelativeFolder,
    string FullPath,
    string RelativePath,
    bool ExistsInDatabase,
    string DatabaseSourcePath);
