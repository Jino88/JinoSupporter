using System.Globalization;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// Housekeeping for the team memo, which lives in the WorkHub memo DB:
/// <list type="bullet">
/// <item>once, copies the memos that were written on the web while it still had its own tables
/// (process-review.db TeamMemo*) into the shared DB, after taking a backup of that DB;</item>
/// <item>every few seconds, notices saves made in WorkHub and tells open pages to reload;</item>
/// <item>keeps every memo pre-translated into English, Korean and Vietnamese: once at start and again
/// shortly after any change (web or WorkHub), so switching language never waits. Only text that is not
/// translated yet is sent. Off when Setting "TeamMemo:AutoTranslate" is 0.</item>
/// </list>
/// </summary>
public sealed class TeamMemoBackgroundService(
    WebRepository repo, TeamMemoService memos, IServiceScopeFactory scopes, ILogger<TeamMemoBackgroundService> log)
    : BackgroundService
{
    private const string MigratedKey = "TeamMemo:MigratedToMemoDb";
    public const string AutoTranslateKey = "TeamMemo:AutoTranslate";
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(5);
    // Wait for edits to settle so a burst of saves costs one pass.
    private static readonly TimeSpan TranslateQuietPeriod = TimeSpan.FromSeconds(20);

    private DateTime? _translateDirtyAt = DateTime.Now;

    private void MarkTranslateDirty(string? _) => _translateDirtyAt = DateTime.Now;

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        try { await Task.Delay(TimeSpan.FromSeconds(5), stoppingToken); } catch (OperationCanceledException) { return; }

        try
        {
            if (repo.GetSetting(MigratedKey) is null && File.Exists(memos.DbPath))
                MigrateLegacyWebMemos();
        }
        catch (Exception ex)
        {
            log.LogError(ex, "Moving the old web team memos into the memo DB failed; will retry on next start");
        }

        memos.Changed += MarkTranslateDirty;
        try
        {
            string? last = null;
            while (!stoppingToken.IsCancellationRequested)
            {
                string? now = memos.ReadFingerprint();
                if (now is not null && last is not null && now != last) memos.RaiseChanged(null);
                if (now is not null) last = now;

                if (_translateDirtyAt is { } dirty && DateTime.Now - dirty > TranslateQuietPeriod
                    && repo.GetSetting(AutoTranslateKey) != "0" && !TeamMemoTranslationService.AllRunning)
                {
                    _translateDirtyAt = null;
                    await PreTranslateAsync(stoppingToken);
                }

                try { await Task.Delay(PollInterval, stoppingToken); } catch (OperationCanceledException) { return; }
            }
        }
        finally
        {
            memos.Changed -= MarkTranslateDirty;
        }
    }

    private async Task PreTranslateAsync(CancellationToken ct)
    {
        try
        {
            using IServiceScope scope = scopes.CreateScope();
            var translations = scope.ServiceProvider.GetRequiredService<TeamMemoTranslationService>();
            if (!translations.IsConfigured || !File.Exists(memos.DbPath)) return;
            int sent = await translations.TranslateAllMemosAsync(ct);
            if (sent > 0) log.LogInformation("Pre-translated {Count} new team memo texts", sent);
        }
        catch (Exception ex)
        {
            log.LogWarning(ex, "Team memo pre-translation failed; will retry after the next change");
        }
    }

    /// <summary>
    /// Copies memos written on the web before 2026-10-03 (Source = '' rows of the old TeamMemo* tables in
    /// process-review.db) into the memo DB. Rows mirrored from WorkHub are skipped — they are already
    /// there. The old tables are left untouched as a record.
    /// </summary>
    public void MigrateLegacyWebMemos()
    {
        string webDb = repo.GetDbPath();
        if (!LegacyTablesExist(webDb))
        {
            repo.SetSetting(MigratedKey, "no legacy tables");
            return;
        }

        string backup = BackupMemoDb();
        log.LogInformation("Memo DB backed up to {Backup} before moving old web memos", backup);

        using var old = new SqliteConnection($"Data Source={webDb};Mode=ReadOnly");
        old.Open();
        bool hasSource = HasColumn(old, "TeamMemos", "Source");
        string local = hasSource ? "AND IFNULL(Source, '') = ''" : "";

        // Folders: re-create by path so an existing folder with the same name is reused.
        var folderMap = new Dictionary<long, string>();
        var oldFolders = new List<(long Id, long? Parent, string Name)>();
        using (var c = old.CreateCommand())
        {
            c.CommandText = $"SELECT Id, ParentId, Name FROM TeamMemoFolders WHERE IsDeleted = 0 {(HasColumn(old, "TeamMemoFolders", "Source") ? "AND IFNULL(Source, '') = ''" : "")} ORDER BY Id;";
            using var r = c.ExecuteReader();
            while (r.Read()) oldFolders.Add((r.GetInt64(0), r.IsDBNull(1) ? null : r.GetInt64(1), r.GetString(2)));
        }

        string? FolderFor(long? oldId)
        {
            if (oldId is not { } fid) return null;
            if (folderMap.TryGetValue(fid, out var mapped)) return mapped;
            var f = oldFolders.FirstOrDefault(x => x.Id == fid);
            if (f.Name is null) return null;
            string? parent = FolderFor(f.Parent);
            string? existing = memos.GetFolders().FirstOrDefault(x => x.ParentId == parent && string.Equals(x.Name, f.Name, StringComparison.OrdinalIgnoreCase))?.Id;
            string id = existing ?? memos.AddFolder(parent, f.Name, "migration");
            folderMap[fid] = id;
            return id;
        }

        int moved = 0;
        var oldMemos = new List<(long Id, long? Folder, string Title, string Kind, string Status, string Owner, string Tags, string? Start, string? Due,
            string Description, bool Pinned, string CreatedBy, string? AiSummary, string? AiAt, string? AiBy, string? AiModel)>();
        using (var c = old.CreateCommand())
        {
            bool hasAi = HasColumn(old, "TeamMemos", "AiSummary");
            c.CommandText = $"""
                SELECT Id, FolderId, Title, Kind, Status, Owner, Tags, StartDate, DueDate, Description, IsPinned, CreatedBy,
                       {(hasAi ? "AiSummary, AiSummaryAt, AiSummaryBy, AiSummaryModel" : "NULL, NULL, NULL, NULL")}
                  FROM TeamMemos WHERE IsDeleted = 0 {local} ORDER BY Id;
                """;
            using var r = c.ExecuteReader();
            while (r.Read())
            {
                string? S(int i) => r.IsDBNull(i) ? null : r.GetString(i);
                oldMemos.Add((r.GetInt64(0), r.IsDBNull(1) ? null : r.GetInt64(1), r.GetString(2), r.GetString(3), r.GetString(4),
                    r.GetString(5), r.GetString(6), S(7), S(8), r.GetString(9), r.GetInt64(10) != 0, r.GetString(11),
                    S(12), S(13), S(14), S(15)));
            }
        }

        foreach (var m in oldMemos)
        {
            string author = m.CreatedBy.Length > 0 ? m.CreatedBy : "migration";
            string kind = TeamMemoKind.All.Contains(m.Kind) ? m.Kind : TeamMemoKind.Task;
            string memoId = memos.CreateMemo(FolderFor(m.Folder), m.Title, kind, author);
            TeamMemo created = memos.GetMemo(memoId)!;
            created.Status = TeamMemoStatus.All.Contains(m.Status) ? m.Status : TeamMemoStatus.Todo;
            created.Owner = m.Owner;
            created.Tags = m.Tags;
            created.StartDate = D(m.Start);
            created.DueDate = D(m.Due);
            created.Description = m.Description;
            created.IsPinned = m.Pinned;
            memos.UpdateMemo(created, author);

            // Entries with their pictures, then entry comments and summaries.
            var entryKeys = new Dictionary<long, string>();
            var entries = new List<(long Id, string Date, string Title, string Body, string? Target, string? Done, string Author, string? Ai, string? AiBy, string? AiModel)>();
            using (var c = old.CreateCommand())
            {
                bool hasAi = HasColumn(old, "TeamMemoEntries", "AiSummary");
                c.CommandText = $"""
                    SELECT Id, EntryDate, Title, Body, TargetDate, DoneDate, Author,
                           {(hasAi ? "AiSummary, AiSummaryBy, AiSummaryModel" : "NULL, NULL, NULL")}
                      FROM TeamMemoEntries WHERE MemoId = @m AND IsDeleted = 0 ORDER BY EntryDate, Id;
                    """;
                c.Parameters.AddWithValue("@m", m.Id);
                using var r = c.ExecuteReader();
                while (r.Read())
                    entries.Add((r.GetInt64(0), r.GetString(1), r.GetString(2), r.GetString(3), r.IsDBNull(4) ? null : r.GetString(4),
                        r.IsDBNull(5) ? null : r.GetString(5), r.GetString(6), r.IsDBNull(7) ? null : r.GetString(7),
                        r.IsDBNull(8) ? null : r.GetString(8), r.IsDBNull(9) ? null : r.GetString(9)));
            }
            foreach (var e in entries)
            {
                var files = ReadFiles(old, "EntryId = @x AND CommentId IS NULL", e.Id);
                string key = memos.AddEntry(memoId, D(e.Date) ?? DateOnly.FromDateTime(DateTime.Today), e.Title, e.Body, D(e.Target),
                    e.Author.Length > 0 ? e.Author : author, files);
                if (e.Done is not null) memos.SetEntryDone(memoId, key, D(e.Done), e.Author.Length > 0 ? e.Author : author);
                if (!string.IsNullOrWhiteSpace(e.Ai)) memos.SaveEntryAiSummary(memoId, key, e.Ai!, e.AiModel ?? "", e.AiBy ?? author);
                entryKeys[e.Id] = key;
            }

            using (var c = old.CreateCommand())
            {
                c.CommandText = "SELECT Id, EntryId, Author, Body, IsResult FROM TeamMemoComments WHERE MemoId = @m AND IsDeleted = 0 ORDER BY CreatedAt, Id;";
                c.Parameters.AddWithValue("@m", m.Id);
                var comments = new List<(long Id, long? Entry, string Author, string Body, bool Result)>();
                using (var r = c.ExecuteReader())
                    while (r.Read()) comments.Add((r.GetInt64(0), r.IsDBNull(1) ? null : r.GetInt64(1), r.GetString(2), r.GetString(3), r.GetInt64(4) != 0));
                foreach (var cm in comments)
                    memos.AddComment(memoId, cm.Entry is { } eid && entryKeys.TryGetValue(eid, out var k) ? k : null, cm.Body, cm.Result, cm.Author,
                        ReadFiles(old, "CommentId = @x", cm.Id));
            }

            if (!string.IsNullOrWhiteSpace(m.AiSummary))
                memos.SaveAiSummary(memoId, m.AiSummary!, m.AiModel ?? "", m.AiBy ?? author);
            moved++;
        }

        repo.SetSetting(MigratedKey, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} moved {moved} memos; backup {backup}");
        log.LogInformation("Moved {Count} old web team memos into the memo DB", moved);
    }

    private static List<(string Name, string ContentType, byte[] Data)> ReadFiles(SqliteConnection old, string where, long id)
    {
        var list = new List<(string, string, byte[])>();
        using var c = old.CreateCommand();
        c.CommandText = $"SELECT FileName, ContentType, Data FROM TeamMemoAttachments WHERE IsDeleted = 0 AND {where} ORDER BY Id;";
        c.Parameters.AddWithValue("@x", id);
        using var r = c.ExecuteReader();
        while (r.Read()) list.Add((r.GetString(0), r.GetString(1), (byte[])r.GetValue(2)));
        return list;
    }

    private string BackupMemoDb()
    {
        string src = memos.DbPath;
        string dest = Path.Combine(Path.GetDirectoryName(src) ?? "", $"jinomemo.backup-{DateTime.Now:yyyyMMdd-HHmmss}.db");
        using var from = new SqliteConnection($"Data Source={src};Mode=ReadOnly");
        using var to = new SqliteConnection($"Data Source={dest}");
        from.Open();
        to.Open();
        from.BackupDatabase(to);
        return dest;
    }

    private static bool LegacyTablesExist(string path)
    {
        if (!File.Exists(path)) return false;
        using var c = new SqliteConnection($"Data Source={path};Mode=ReadOnly");
        c.Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = 'TeamMemos';";
        return Convert.ToInt64(cmd.ExecuteScalar()) == 1;
    }

    private static bool HasColumn(SqliteConnection c, string table, string column)
    {
        using var cmd = c.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info({table});";
        using var r = cmd.ExecuteReader();
        while (r.Read()) if (string.Equals(r.GetString(1), column, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static DateOnly? D(string? s) =>
        DateOnly.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;
}
