using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

public static class TeamMemoStatus
{
    public const string Todo       = "Todo";
    public const string InProgress = "InProgress";
    public const string Hold       = "Hold";
    public const string Done       = "Done";
    public const string Canceled   = "Canceled";

    public static readonly string[] All = [Todo, InProgress, Hold, Done, Canceled];

    public static string Label(string status) => status switch
    {
        InProgress => "In progress",
        Hold       => "On hold",
        _          => status,
    };

    /// <summary>WorkHub MemoStatus: Todo 0, Doing 1, Done 2, Hold 3, Canceled 4.</summary>
    internal static string FromWorkHub(long v) => v switch
    {
        1 => InProgress, 2 => Done, 3 => Hold, 4 => Canceled, _ => Todo,
    };

    internal static int ToWorkHub(string s) => s switch
    {
        InProgress => 1, Done => 2, Hold => 3, Canceled => 4, _ => 0,
    };

    public static bool IsClosed(string s) => s is Done or Canceled;
}

public static class TeamMemoKind
{
    public const string Task    = "Task";
    public const string Review  = "Review";
    public const string History = "History";

    public static readonly string[] All = [Task, History, Review];

    /// <summary>WorkHub MemoKind: Task 0, Review 1, History 2.</summary>
    internal static string FromWorkHub(long v) => v switch { 1 => Review, 2 => History, _ => Task };
    internal static int ToWorkHub(string s) => s switch { Review => 1, History => 2, _ => 0 };
}

public sealed class TeamMemoFolder
{
    public string  Id       { get; set; } = "";
    public string? ParentId { get; set; }
    public string  Name     { get; set; } = "";
    public int     Ordinal  { get; set; }
    public string  Path     { get; set; } = "";
    /// <summary>Hidden in WorkHub's work-memo explorer (e.g. meeting notes).</summary>
    public bool    IsHidden { get; set; }
    /// <summary>Web-only: a Plan folder — shown under Plan and kept out of Work Memo and Board.</summary>
    public bool    IsPlan   { get; set; }
}

/// <summary>One memo as shown in the tree, board and schedule — no entries or comments.</summary>
public sealed class TeamMemoSummary
{
    public string    Id            { get; set; } = "";
    public string?   FolderId      { get; set; }
    public string    Title         { get; set; } = "";
    public string    Kind          { get; set; } = TeamMemoKind.Task;
    public string    Status        { get; set; } = TeamMemoStatus.Todo;
    public string    Owner         { get; set; } = "";
    public string    Tags          { get; set; } = "";
    public DateOnly? StartDate     { get; set; }
    public DateOnly? DueDate       { get; set; }
    public bool      IsPinned      { get; set; }
    public string    UpdatedBy     { get; set; } = "";
    public DateTime  UpdatedAt     { get; set; }
    public int       EntryCount    { get; set; }
    public int       CommentCount  { get; set; }
    public List<string> Participants { get; set; } = new();
    public string?   LastCommentBy   { get; set; }
    public string?   LastCommentBody { get; set; }
    public bool      LastCommentIsResult { get; set; }
    public DateTime? LastActivityAt  { get; set; }
    public DateTime  CreatedAt       { get; set; }
    /// <summary>WorkHub weekly repeat: DayOfWeek numbers (0 = Sunday) joined by commas; "" = no repeat.</summary>
    public string    RepeatWeekdays  { get; set; } = "";
    public DateOnly? RepeatUntil     { get; set; }
    public DateTime? ScheduleStartAt { get; set; }
    public DateTime? ScheduleEndAt   { get; set; }
    public bool      ScheduleAllDay  { get; set; }

    public bool IsOverdue(DateOnly today) =>
        !TeamMemoStatus.IsClosed(Status) && DueDate is { } due && due < today;
}

public sealed class TeamMemo
{
    public string    Id          { get; set; } = "";
    public string?   FolderId    { get; set; }
    public string    Title       { get; set; } = "";
    public string    Kind        { get; set; } = TeamMemoKind.Task;
    public string    Status      { get; set; } = TeamMemoStatus.Todo;
    public string    Owner       { get; set; } = "";
    public string    Tags        { get; set; } = "";
    public DateOnly? StartDate   { get; set; }
    public DateOnly? DueDate     { get; set; }
    /// <summary>Web-side summary text (WorkHub has no such field; History memos show their fields here).</summary>
    public string    Description { get; set; } = "";
    public bool      IsPinned    { get; set; }
    public string    CreatedBy   { get; set; } = "";
    public DateTime  CreatedAt   { get; set; }
    public string    UpdatedBy   { get; set; } = "";
    public DateTime  UpdatedAt   { get; set; }
    /// <summary>WorkHub's optimistic-concurrency revision; every save from either app bumps it.</summary>
    public int       Revision    { get; set; }

    /// <summary>WorkHub weekly repeat: DayOfWeek numbers (0 = Sunday) joined by commas; "" = no repeat.</summary>
    public string    RepeatWeekdays { get; set; } = "";
    public DateOnly? RepeatUntil    { get; set; }

    public string?   AiSummary      { get; set; }
    public DateTime? AiSummaryAt    { get; set; }
    public string    AiSummaryBy    { get; set; } = "";
    public string    AiSummaryModel { get; set; } = "";
    public int?      AiSummaryRevision { get; set; }

    public List<TeamMemoEntry>      Entries     { get; set; } = new();
    /// <summary>Comments on the memo itself rather than on one entry.</summary>
    public List<TeamMemoComment>    Comments    { get; set; } = new();
    /// <summary>Files attached to memo-level comments (history pictures hang off their entries).</summary>
    public List<TeamMemoAttachment> Attachments { get; set; } = new();

    /// <summary>The memo or any comment changed after the saved AI summary.</summary>
    public bool IsAiSummaryStale =>
        AiSummaryAt is { } at && (AiSummaryRevision != Revision || NewestCommentAt > at);

    private DateTime NewestCommentAt =>
        Comments.Concat(Entries.SelectMany(e => e.Comments)).Select(c => c.UpdatedAt).DefaultIfEmpty(DateTime.MinValue).Max();
}

/// <summary>A dated block of the memo body: <c>[yyyy-MM-dd HH:mm|target|done|dN|c]</c> + title + text.</summary>
public sealed class TeamMemoEntry
{
    /// <summary>Stable key inside the memo: the block timestamp, plus <c>#n</c> for repeats.</summary>
    public string    Id         { get; set; } = "";
    public string    MemoId     { get; set; } = "";
    public DateOnly  EntryDate  { get; set; }
    public string    Title      { get; set; } = "";
    public string    Body       { get; set; } = "";
    public DateOnly? TargetDate { get; set; }
    public DateOnly? DoneDate   { get; set; }
    /// <summary>WorkHub sub-entry depth (0 = top level).</summary>
    public int       Depth      { get; set; }
    public string    Author     { get; set; } = "";
    public DateTime  CreatedAt  { get; set; }
    public DateTime  UpdatedAt  { get; set; }

    public string?   AiSummary      { get; set; }
    public DateTime? AiSummaryAt    { get; set; }
    public string    AiSummaryBy    { get; set; } = "";
    public string    AiSummaryModel { get; set; } = "";
    internal string? AiSummaryHash  { get; set; }
    internal string  ContentHash    { get; set; } = "";

    public List<TeamMemoComment>    Comments    { get; set; } = new();
    public List<TeamMemoAttachment> Attachments { get; set; } = new();

    public bool IsDone => DoneDate is not null;

    /// <summary>The entry text/pictures changed, or a comment came in, after its AI summary.</summary>
    public bool IsAiSummaryStale =>
        AiSummaryAt is { } at
        && (AiSummaryHash != ContentHash
            || Comments.Any(c => c.UpdatedAt > at || c.Attachments.Any(f => f.CreatedAt > at)));
}

public sealed class TeamMemoComment
{
    public long     Id        { get; set; }
    public string   MemoId    { get; set; } = "";
    public string?  EntryId   { get; set; }
    public string   Author    { get; set; } = "";
    public string   Body      { get; set; } = "";
    public bool     IsResult  { get; set; }
    public DateTime CreatedAt { get; set; }
    public DateTime UpdatedAt { get; set; }

    public List<TeamMemoAttachment> Attachments { get; set; } = new();
}

public sealed class TeamMemoAttachment
{
    /// <summary><c>w-</c> + WorkHub attachment id for history pictures, <c>c-</c> + number for comment files.</summary>
    public string   Id          { get; set; } = "";
    public string   MemoId      { get; set; } = "";
    public string?  EntryId     { get; set; }
    public long?    CommentId   { get; set; }
    public string   FileName    { get; set; } = "";
    public string   ContentType { get; set; } = "";
    public long     Size        { get; set; }
    public string   Author      { get; set; } = "";
    public DateTime CreatedAt   { get; set; }

    public bool IsImage => ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
}

/// <summary>Someone else saved the memo after it was loaded; reload and retry.</summary>
public sealed class TeamMemoConflictException(string message) : InvalidOperationException(message);

/// <summary>
/// Team memo stored in the WorkHub memo DB (jinomemo.db), so the web and WorkHub read and write one
/// database. Folders are WorkHub's MemoCategories, memos its Memos rows, and history entries the
/// dated blocks inside Memo.Content — written in WorkHub's own block format. Every write follows
/// WorkHub's rules: Revision + UpdatedAt bump, Category path copy, Tags/MemoTags, the MemoSearch
/// FTS row, and attachments as Attachments + AttachmentBlobs + the file under the attachments root.
///
/// What WorkHub has no place for — who wrote what, comments, results, AI summaries — lives in
/// Web* tables in the same file; WorkHub's EF model never touches them.
/// </summary>
public sealed class TeamMemoService
{
    public const long MaxAttachmentBytes = 20L * 1024 * 1024;
    public const string DefaultDbPath = @"D:\000. MyWorks\002. DB\WorkHub\MemoDb\jinomemo.db";
    public const int MaxFolderDepth = 8;
    private const int ImageWidth = 440;

    private readonly WebRepository _repo;
    private readonly object _schemaGate = new();
    private string? _schemaReadyFor;

    public TeamMemoService(WebRepository repo) => _repo = repo;

    /// <summary>Raised after any write (memo id, or null for tree-wide changes) so open pages reload.</summary>
    public event Action<string?>? Changed;

    internal void RaiseChanged(string? memoId)
    {
        try { Changed?.Invoke(memoId); } catch { /* a listener failing must not fail the write */ }
    }

    public string DbPath => _repo.GetSetting("TeamMemo:MemoDbPath") is { Length: > 0 } p ? p : DefaultDbPath;

    public string AttachmentsRoot => _repo.GetSetting("TeamMemo:MemoAttachmentsDir") is { Length: > 0 } a
        ? a
        : Path.Combine(Path.GetDirectoryName(DbPath) ?? "", "attachments");

    /// <summary>Name shown for content written in WorkHub, which records no author.</summary>
    public string WorkHubAuthor => _repo.GetSetting("TeamMemo:WorkHubAuthor") is { Length: > 0 } a ? a : "WorkHub";

    public void SaveStoreSettings(string dbPath, string attachmentsDir, string workHubAuthor)
    {
        _repo.SetSetting("TeamMemo:MemoDbPath", dbPath.Trim());
        _repo.SetSetting("TeamMemo:MemoAttachmentsDir", attachmentsDir.Trim());
        _repo.SetSetting("TeamMemo:WorkHubAuthor", workHubAuthor.Trim());
        lock (_schemaGate) _schemaReadyFor = null;
        RaiseChanged(null);
    }

    internal SqliteConnection Open()
    {
        string path = DbPath;
        if (!File.Exists(path))
            throw new InvalidOperationException($"Memo DB not found: {path}. Open WorkHub once or fix the path in Setting.");
        var conn = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode = SqliteOpenMode.ReadWrite,
            DefaultTimeout = 30,
        }.ToString());
        conn.Open();
        using (var pragma = conn.CreateCommand())
        {
            pragma.CommandText = "PRAGMA foreign_keys = ON; PRAGMA busy_timeout = 15000;";
            pragma.ExecuteNonQuery();
        }
        lock (_schemaGate)
        {
            if (_schemaReadyFor != path)
            {
                EnsureWebSchema(conn);
                _schemaReadyFor = path;
            }
        }
        return conn;
    }

    private static void EnsureWebSchema(SqliteConnection conn)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS WebMemoMeta (
                MemoId            TEXT PRIMARY KEY,
                Owner             TEXT NOT NULL DEFAULT '',
                Description       TEXT NOT NULL DEFAULT '',
                CreatedBy         TEXT NOT NULL DEFAULT '',
                UpdatedBy         TEXT NOT NULL DEFAULT '',
                AiSummary         TEXT NULL,
                AiSummaryAt       TEXT NULL,
                AiSummaryBy       TEXT NOT NULL DEFAULT '',
                AiSummaryModel    TEXT NOT NULL DEFAULT '',
                AiSummaryRevision INTEGER NULL
            );
            CREATE TABLE IF NOT EXISTS WebEntryMeta (
                MemoId         TEXT NOT NULL,
                EntryKey       TEXT NOT NULL,
                Author         TEXT NOT NULL DEFAULT '',
                UpdatedAt      TEXT NULL,
                AiSummary      TEXT NULL,
                AiSummaryAt    TEXT NULL,
                AiSummaryBy    TEXT NOT NULL DEFAULT '',
                AiSummaryModel TEXT NOT NULL DEFAULT '',
                AiSummaryHash  TEXT NULL,
                PRIMARY KEY (MemoId, EntryKey)
            );
            CREATE TABLE IF NOT EXISTS WebMemoComments (
                Id        INTEGER PRIMARY KEY AUTOINCREMENT,
                MemoId    TEXT NOT NULL,
                EntryKey  TEXT NULL,
                Author    TEXT NOT NULL DEFAULT '',
                Body      TEXT NOT NULL DEFAULT '',
                IsResult  INTEGER NOT NULL DEFAULT 0,
                CreatedAt TEXT NOT NULL,
                UpdatedAt TEXT NOT NULL,
                IsDeleted INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS IX_WebMemoComments_Memo ON WebMemoComments(MemoId);
            CREATE TABLE IF NOT EXISTS WebCommentFiles (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                MemoId      TEXT NOT NULL,
                CommentId   INTEGER NOT NULL,
                FileName    TEXT NOT NULL,
                ContentType TEXT NOT NULL,
                Size        INTEGER NOT NULL,
                Data        BLOB NOT NULL,
                Author      TEXT NOT NULL DEFAULT '',
                CreatedAt   TEXT NOT NULL,
                IsDeleted   INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS IX_WebCommentFiles_Memo ON WebCommentFiles(MemoId);
            CREATE TABLE IF NOT EXISTS WebFolderMeta (
                FolderId TEXT PRIMARY KEY,
                IsPlan   INTEGER NOT NULL DEFAULT 0
            );
            """;
        cmd.ExecuteNonQuery();
    }

    /// <summary>Cheap fingerprint of the WorkHub tables, polled to notice saves made in WorkHub.</summary>
    internal string? ReadFingerprint()
    {
        try
        {
            using SqliteConnection conn = Open();
            using SqliteCommand cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT (SELECT COUNT(*) || ':' || IFNULL(SUM(Revision), 0) || ':' || IFNULL(MAX(UpdatedAt), 0) FROM Memos)
                       || '|' || (SELECT COUNT(*) || ':' || IFNULL(group_concat(Path, '/'), '') FROM MemoCategories);
                """;
            return cmd.ExecuteScalar() as string;
        }
        catch
        {
            return null;
        }
    }

    // ── Folders (MemoCategories) ───────────────────────────────────────────────

    public List<TeamMemoFolder> GetFolders()
    {
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT c.Id, c.ParentId, c.Name, c.Ordinal, c.Path, c.IsHidden, IFNULL(w.IsPlan, 0)
            FROM MemoCategories c LEFT JOIN WebFolderMeta w ON w.FolderId = c.Id
            ORDER BY c.Ordinal, c.Name;
            """;
        var list = new List<TeamMemoFolder>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new TeamMemoFolder
            {
                Id = r.GetString(0),
                ParentId = r.IsDBNull(1) ? null : r.GetString(1),
                Name = r.GetString(2),
                Ordinal = r.GetInt32(3),
                Path = r.GetString(4),
                IsHidden = r.GetInt64(5) != 0,
                IsPlan = r.GetInt64(6) != 0,
            });
        return list;
    }

    public string AddFolder(string? parentId, string name, string user)
    {
        name = ValidateFolderName(name);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        var all = LoadCategoryRows(conn, tx);
        string parentPath = "";
        if (parentId is not null)
        {
            if (!all.TryGetValue(parentId, out var parent)) throw new InvalidOperationException("The parent folder no longer exists.");
            if (DepthOf(all, parentId) + 1 >= MaxFolderDepth)
                throw new InvalidOperationException($"Folders can be nested at most {MaxFolderDepth} levels.");
            parentPath = parent.Path;
        }
        var siblings = all.Values.Where(c => c.ParentId == parentId).ToList();
        if (siblings.Any(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException($"'{name}' already exists here.");

        string id = NewId();
        Exec(conn, tx, """
            INSERT INTO MemoCategories (Id, ParentId, Name, Path, Ordinal, CreatedAt, IsHidden)
            VALUES (@id, @p, @n, @path, @o, @t, 0);
            """, ("@id", id), ("@p", parentId), ("@n", name),
            ("@path", parentPath.Length == 0 ? name : parentPath + "/" + name), ("@o", siblings.Count), ("@t", Bin(DateTimeOffset.Now)));
        tx.Commit();
        RaiseChanged(null);
        return id;
    }

    public void RenameFolder(string id, string name)
    {
        name = ValidateFolderName(name);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        var all = LoadCategoryRows(conn, tx);
        if (!all.TryGetValue(id, out var node)) throw new InvalidOperationException("That folder no longer exists.");
        if (node.Name == name) return;
        if (all.Values.Any(c => c.ParentId == node.ParentId && c.Id != id && c.Name == name))
            throw new InvalidOperationException($"'{name}' already exists here.");
        Exec(conn, tx, "UPDATE MemoCategories SET Name = @n WHERE Id = @id;", ("@n", name), ("@id", id));
        node.Name = name;
        RewritePaths(conn, tx, all, id);
        tx.Commit();
        RaiseChanged(null);
    }

    /// <summary>Puts a folder (with everything under it) inside another, or at the top when null.</summary>
    public void MoveFolder(string id, string? newParentId)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        var all = LoadCategoryRows(conn, tx);
        if (!all.TryGetValue(id, out var node)) throw new InvalidOperationException("That folder no longer exists.");
        if (newParentId is not null)
        {
            if (!all.ContainsKey(newParentId)) throw new InvalidOperationException("The target folder no longer exists.");
            for (string? cur = newParentId; cur is not null; cur = all.TryGetValue(cur, out var c) ? c.ParentId : null)
                if (cur == id) throw new InvalidOperationException($"'{node.Name}' cannot be moved into itself or its own sub-folder.");
        }
        if (node.ParentId == newParentId) return;
        var siblings = all.Values.Where(c => c.ParentId == newParentId && c.Id != id).ToList();
        if (siblings.Any(c => c.Name == node.Name))
            throw new InvalidOperationException($"'{node.Name}' already exists there.");
        if (newParentId is not null && DepthOf(all, newParentId) + 1 + HeightOf(all, id) > MaxFolderDepth)
            throw new InvalidOperationException($"Folders can be nested at most {MaxFolderDepth} levels.");

        Exec(conn, tx, "UPDATE MemoCategories SET ParentId = @p, Ordinal = @o WHERE Id = @id;",
            ("@p", newParentId), ("@o", siblings.Count), ("@id", id));
        node.ParentId = newParentId;
        RewritePaths(conn, tx, all, id);
        tx.Commit();
        RaiseChanged(null);
    }

    /// <summary>
    /// Marks a folder as a meeting folder or back. This is WorkHub's "hidden" flag: such folders (and
    /// everything under them) are left out of Work Memo and Board and listed under Meetings instead —
    /// WorkHub likewise keeps them out of its work-memo explorer.
    /// </summary>
    public void SetFolderHidden(string id, bool hidden)
    {
        Execute("UPDATE MemoCategories SET IsHidden = @h WHERE Id = @id;", ("@h", hidden ? 1 : 0), ("@id", id));
        RaiseChanged(null);
    }

    /// <summary>Marks a folder as a Plan folder (web-only flag) or back to a normal work-memo folder.</summary>
    public void SetFolderPlan(string id, bool plan)
    {
        Execute("""
            INSERT INTO WebFolderMeta (FolderId, IsPlan) VALUES (@id, @p)
            ON CONFLICT(FolderId) DO UPDATE SET IsPlan = excluded.IsPlan;
            """, ("@id", id), ("@p", plan ? 1 : 0));
        RaiseChanged(null);
    }

    /// <summary>Ids of Plan folders and every folder below them.</summary>
    public static HashSet<string> PlanFolderIds(IReadOnlyCollection<TeamMemoFolder> folders) =>
        FlaggedFolderIds(folders, f => f.IsPlan);

    /// <summary>Ids of meeting folders and every folder below them.</summary>
    public static HashSet<string> MeetingFolderIds(IReadOnlyCollection<TeamMemoFolder> folders) =>
        FlaggedFolderIds(folders, f => f.IsHidden);

    private static HashSet<string> FlaggedFolderIds(IReadOnlyCollection<TeamMemoFolder> folders, Func<TeamMemoFolder, bool> flag)
    {
        var byId = folders.ToDictionary(f => f.Id, StringComparer.Ordinal);
        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (TeamMemoFolder f in folders)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (TeamMemoFolder? cur = f; cur is not null && seen.Add(cur.Id);
                 cur = cur.ParentId is { } p && byId.TryGetValue(p, out var parent) ? parent : null)
            {
                if (flag(cur)) { result.Add(f.Id); break; }
            }
        }
        return result;
    }

    /// <summary>
    /// Deletes a folder the WorkHub way: its sub-folders and memos are not deleted but move up one level.
    /// </summary>
    public void DeleteFolder(string id)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        var all = LoadCategoryRows(conn, tx);
        if (!all.TryGetValue(id, out var node)) return;
        foreach (var child in all.Values.Where(c => c.ParentId == id).ToList())
        {
            Exec(conn, tx, "UPDATE MemoCategories SET ParentId = @p WHERE Id = @id;", ("@p", node.ParentId), ("@id", child.Id));
            child.ParentId = node.ParentId;
            RewritePaths(conn, tx, all, child.Id);
        }
        string parentPath = node.ParentId is { } p && all.TryGetValue(p, out var parent) ? parent.Path : "";
        Exec(conn, tx, "UPDATE Memos SET CategoryId = @p, Category = @path WHERE CategoryId = @id;",
            ("@p", node.ParentId), ("@path", parentPath), ("@id", id));
        Exec(conn, tx, "DELETE FROM MemoCategories WHERE Id = @id;", ("@id", id));
        tx.Commit();
        RaiseChanged(null);
    }

    private sealed class CategoryRow
    {
        public string Id = "";
        public string? ParentId;
        public string Name = "";
        public string Path = "";
    }

    private static Dictionary<string, CategoryRow> LoadCategoryRows(SqliteConnection conn, SqliteTransaction tx)
    {
        var all = new Dictionary<string, CategoryRow>(StringComparer.Ordinal);
        using var cmd = Cmd(conn, tx, "SELECT Id, ParentId, Name, Path FROM MemoCategories;");
        using var r = cmd.ExecuteReader();
        while (r.Read())
            all[r.GetString(0)] = new CategoryRow
            {
                Id = r.GetString(0), ParentId = r.IsDBNull(1) ? null : r.GetString(1), Name = r.GetString(2), Path = r.GetString(3),
            };
        return all;
    }

    /// <summary>Rewrites the path of a folder, every descendant, and the Category copy on their memos.</summary>
    private static void RewritePaths(SqliteConnection conn, SqliteTransaction tx, Dictionary<string, CategoryRow> all, string id)
    {
        CategoryRow node = all[id];
        string parentPath = node.ParentId is { } p && all.TryGetValue(p, out var parent) ? parent.Path : "";
        void Walk(CategoryRow n, string basePath, int guard)
        {
            if (guard > MaxFolderDepth + 2) return;
            n.Path = basePath.Length == 0 ? n.Name : basePath + "/" + n.Name;
            Exec(conn, tx, "UPDATE MemoCategories SET Path = @path WHERE Id = @id;", ("@path", n.Path), ("@id", n.Id));
            Exec(conn, tx, "UPDATE Memos SET Category = @path WHERE CategoryId = @id;", ("@path", n.Path), ("@id", n.Id));
            foreach (var child in all.Values.Where(c => c.ParentId == n.Id).ToList()) Walk(child, n.Path, guard + 1);
        }
        Walk(node, parentPath, 0);
    }

    private static int DepthOf(Dictionary<string, CategoryRow> all, string id)
    {
        int depth = 0;
        for (string? cur = all.TryGetValue(id, out var n) ? n.ParentId : null; cur is not null && depth < MaxFolderDepth + 2;
             cur = all.TryGetValue(cur, out var c) ? c.ParentId : null)
            depth++;
        return depth;
    }

    private static int HeightOf(Dictionary<string, CategoryRow> all, string id)
    {
        var children = all.Values.Where(c => c.ParentId == id).ToList();
        return children.Count == 0 ? 0 : 1 + children.Max(c => HeightOf(all, c.Id));
    }

    private static string ValidateFolderName(string name)
    {
        string trimmed = (name ?? "").Trim();
        if (trimmed.Length == 0) throw new InvalidOperationException("Folder name is required.");
        if (trimmed.Contains('/')) throw new InvalidOperationException("Folder names cannot contain '/'.");
        return trimmed.Length > 60 ? trimmed[..60] : trimmed;
    }

    // ── Memos: read ────────────────────────────────────────────────────────────

    public List<TeamMemoSummary> ListMemos()
    {
        using SqliteConnection conn = Open();
        string wh = WorkHubAuthor;
        var list = new List<TeamMemoSummary>();
        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT m.Id, m.CategoryId, m.Title, m.Kind, m.Status, m.Tags, m.StartDate, m.DueDate, m.IsPinned,
                       m.UpdatedAt, m.Content, IFNULL(w.Owner, ''), IFNULL(w.UpdatedBy, ''), IFNULL(w.CreatedBy, ''),
                       (SELECT COUNT(*) FROM WebMemoComments c WHERE c.MemoId = m.Id AND c.IsDeleted = 0),
                       lc.Author, lc.Body, lc.IsResult, lc.CreatedAt,
                       m.CreatedAt, m.RepeatWeekdays, m.RepeatUntil, m.ScheduleStartAt, m.ScheduleEndAt, m.ScheduleAllDay
                  FROM Memos m
                  LEFT JOIN WebMemoMeta w ON w.MemoId = m.Id
                  LEFT JOIN WebMemoComments lc ON lc.Id = (
                        SELECT c.Id FROM WebMemoComments c WHERE c.MemoId = m.Id AND c.IsDeleted = 0
                         ORDER BY c.CreatedAt DESC, c.Id DESC LIMIT 1)
                 ORDER BY m.IsPinned DESC, m.UpdatedAt DESC;
                """;
            using SqliteDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                DateTime updated = FromBin(r.GetInt64(9));
                DateTime? lastComment = r.IsDBNull(18) ? null : ParseTime(r.GetString(18));
                string owner = r.GetString(11);
                var s = new TeamMemoSummary
                {
                    Id = r.GetString(0),
                    FolderId = r.IsDBNull(1) ? null : r.GetString(1),
                    Title = r.GetString(2),
                    Kind = TeamMemoKind.FromWorkHub(r.GetInt64(3)),
                    Status = TeamMemoStatus.FromWorkHub(r.GetInt64(4)),
                    Tags = r.GetString(5),
                    StartDate = r.IsDBNull(6) ? null : DateOnly.FromDateTime(FromBin(r.GetInt64(6))),
                    DueDate = r.IsDBNull(7) ? null : DateOnly.FromDateTime(FromBin(r.GetInt64(7))),
                    IsPinned = r.GetInt64(8) != 0,
                    UpdatedAt = updated,
                    EntryCount = ContentDoc.Parse(r.GetString(10)).AllBlocks.Count(),
                    Owner = owner.Length > 0 ? owner : wh,
                    UpdatedBy = r.GetString(12) is { Length: > 0 } ub ? ub : wh,
                    CommentCount = r.GetInt32(14),
                    LastCommentBy = r.IsDBNull(15) ? null : r.GetString(15),
                    LastCommentBody = r.IsDBNull(16) ? null : r.GetString(16),
                    LastCommentIsResult = !r.IsDBNull(17) && r.GetInt64(17) != 0,
                    LastActivityAt = lastComment is { } lc && lc > updated ? lc : updated,
                    CreatedAt = FromBin(r.GetInt64(19)),
                    RepeatWeekdays = r.IsDBNull(20) ? "" : r.GetString(20),
                    RepeatUntil = r.IsDBNull(21) ? null : DateOnly.FromDateTime(FromBin(r.GetInt64(21))),
                    ScheduleStartAt = r.IsDBNull(22) ? null : FromBin(r.GetInt64(22)),
                    ScheduleEndAt = r.IsDBNull(23) ? null : FromBin(r.GetInt64(23)),
                    ScheduleAllDay = !r.IsDBNull(24) && r.GetInt64(24) != 0,
                };
                s.Participants.Add(s.Owner);
                if (r.GetString(13) is { Length: > 0 } cb && !s.Participants.Contains(cb, StringComparer.OrdinalIgnoreCase)) s.Participants.Add(cb);
                list.Add(s);
            }
        }

        var byId = list.ToDictionary(m => m.Id, StringComparer.Ordinal);
        using (SqliteCommand who = conn.CreateCommand())
        {
            who.CommandText = """
                SELECT MemoId, Author FROM WebMemoComments WHERE IsDeleted = 0 AND Author <> ''
                UNION ALL SELECT MemoId, Author FROM WebEntryMeta WHERE Author <> '';
                """;
            using SqliteDataReader r = who.ExecuteReader();
            while (r.Read())
                if (byId.TryGetValue(r.GetString(0), out var m) && !m.Participants.Contains(r.GetString(1), StringComparer.OrdinalIgnoreCase))
                    m.Participants.Add(r.GetString(1));
        }
        return list;
    }

    public TeamMemo? GetMemo(string id)
    {
        using SqliteConnection conn = Open();
        string wh = WorkHubAuthor;
        TeamMemo memo;
        string content;
        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT m.Id, m.CategoryId, m.Title, m.Kind, m.Status, m.Tags, m.StartDate, m.DueDate, m.IsPinned,
                       m.CreatedAt, m.UpdatedAt, m.Revision, m.Content,
                       m.HistoryOwner, m.HistoryTrigger, m.HistoryAction, m.HistoryResult, m.ReviewSubject, m.ReviewConclusion,
                       w.Owner, w.Description, w.CreatedBy, w.UpdatedBy, w.AiSummary, w.AiSummaryAt, w.AiSummaryBy,
                       w.AiSummaryModel, w.AiSummaryRevision, m.RepeatWeekdays, m.RepeatUntil
                  FROM Memos m LEFT JOIN WebMemoMeta w ON w.MemoId = m.Id
                 WHERE m.Id = @id;
                """;
            cmd.Parameters.AddWithValue("@id", id);
            using SqliteDataReader r = cmd.ExecuteReader();
            if (!r.Read()) return null;
            string Str(int i) => r.IsDBNull(i) ? "" : r.GetString(i);
            string description = Str(20);
            if (description.Length == 0)
            {
                // WorkHub History / Review memos keep their facts in fields of their own.
                var sb = new StringBuilder();
                void Line(string label, string v) { if (v.Trim().Length > 0) sb.AppendLine($"{label}: {v.Trim()}"); }
                Line("Owner", Str(13)); Line("Trigger", Str(14)); Line("Action", Str(15)); Line("Result", Str(16));
                Line("Review subject", Str(17)); Line("Conclusion", Str(18));
                description = sb.ToString().TrimEnd();
            }
            memo = new TeamMemo
            {
                Id = r.GetString(0),
                FolderId = r.IsDBNull(1) ? null : r.GetString(1),
                Title = r.GetString(2),
                Kind = TeamMemoKind.FromWorkHub(r.GetInt64(3)),
                Status = TeamMemoStatus.FromWorkHub(r.GetInt64(4)),
                Tags = r.GetString(5),
                StartDate = r.IsDBNull(6) ? null : DateOnly.FromDateTime(FromBin(r.GetInt64(6))),
                DueDate = r.IsDBNull(7) ? null : DateOnly.FromDateTime(FromBin(r.GetInt64(7))),
                IsPinned = r.GetInt64(8) != 0,
                CreatedAt = FromBin(r.GetInt64(9)),
                UpdatedAt = FromBin(r.GetInt64(10)),
                Revision = r.GetInt32(11),
                Description = description,
                Owner = Str(19) is { Length: > 0 } o ? o : wh,
                CreatedBy = Str(21) is { Length: > 0 } cb ? cb : wh,
                UpdatedBy = Str(22) is { Length: > 0 } ub ? ub : wh,
                AiSummary = r.IsDBNull(23) ? null : r.GetString(23),
                AiSummaryAt = r.IsDBNull(24) ? null : ParseTime(r.GetString(24)),
                AiSummaryBy = Str(25),
                AiSummaryModel = Str(26),
                AiSummaryRevision = r.IsDBNull(27) ? null : r.GetInt32(27),
                RepeatWeekdays = r.IsDBNull(28) ? "" : r.GetString(28),
                RepeatUntil = r.IsDBNull(29) ? null : DateOnly.FromDateTime(FromBin(r.GetInt64(29))),
            };
            content = r.GetString(12);
        }

        // Entries from the body, then their web-side meta.
        ContentDoc doc = ContentDoc.Parse(content);
        var entries = new Dictionary<string, TeamMemoEntry>(StringComparer.Ordinal);
        var imageEntry = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (ContentBlock b in doc.AllBlocks)
        {
            DateTime when = b.IsLead ? memo.CreatedAt : b.Timestamp;
            var e = new TeamMemoEntry
            {
                Id = b.Key,
                MemoId = id,
                EntryDate = DateOnly.FromDateTime(when),
                Title = b.Title.Length > 0 ? b.Title : "(no title)",
                Body = b.Body,
                TargetDate = b.Target,
                DoneDate = b.Done ? b.Target ?? DateOnly.FromDateTime(when) : null,
                Depth = b.Depth,
                Author = wh,
                CreatedAt = when,
                UpdatedAt = memo.UpdatedAt,
                ContentHash = b.Hash(),
            };
            entries[b.Key] = e;
            memo.Entries.Add(e);
            foreach (var img in b.Images) imageEntry.TryAdd(img.Id, b.Key);
        }

        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT EntryKey, Author, UpdatedAt, AiSummary, AiSummaryAt, AiSummaryBy, AiSummaryModel, AiSummaryHash
                  FROM WebEntryMeta WHERE MemoId = @id;
                """;
            cmd.Parameters.AddWithValue("@id", id);
            using SqliteDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (!entries.TryGetValue(r.GetString(0), out var e)) continue;
                if (r.GetString(1) is { Length: > 0 } a) e.Author = a;
                if (!r.IsDBNull(2)) e.UpdatedAt = ParseTime(r.GetString(2));
                e.AiSummary = r.IsDBNull(3) ? null : r.GetString(3);
                e.AiSummaryAt = r.IsDBNull(4) ? null : ParseTime(r.GetString(4));
                e.AiSummaryBy = r.GetString(5);
                e.AiSummaryModel = r.GetString(6);
                e.AiSummaryHash = r.IsDBNull(7) ? null : r.GetString(7);
            }
        }

        // History pictures: WorkHub attachments referenced from the blocks, in block order.
        if (imageEntry.Count > 0)
        {
            using SqliteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Id, FileName, ContentType, SizeBytes, CreatedAt FROM Attachments WHERE MemoId = @id;";
            cmd.Parameters.AddWithValue("@id", id);
            var atts = new Dictionary<string, TeamMemoAttachment>(StringComparer.Ordinal);
            using (SqliteDataReader r = cmd.ExecuteReader())
            {
                while (r.Read())
                    atts[r.GetString(0)] = new TeamMemoAttachment
                    {
                        Id = "w-" + r.GetString(0),
                        MemoId = id,
                        FileName = r.GetString(1),
                        ContentType = r.GetString(2),
                        Size = r.GetInt64(3),
                        CreatedAt = FromBin(r.GetInt64(4)),
                    };
            }
            foreach (ContentBlock b in doc.AllBlocks)
                foreach (var img in b.Images)
                    if (atts.TryGetValue(img.Id, out var a) && imageEntry[img.Id] == b.Key)
                    {
                        a.EntryId = b.Key;
                        a.Author = entries[b.Key].Author;
                        entries[b.Key].Attachments.Add(a);
                    }
        }

        // Comments and their files.
        var comments = new Dictionary<long, TeamMemoComment>();
        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT Id, EntryKey, Author, Body, IsResult, CreatedAt, UpdatedAt
                  FROM WebMemoComments WHERE MemoId = @id AND IsDeleted = 0 ORDER BY CreatedAt, Id;
                """;
            cmd.Parameters.AddWithValue("@id", id);
            using SqliteDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                var c = new TeamMemoComment
                {
                    Id = r.GetInt64(0),
                    MemoId = id,
                    EntryId = r.IsDBNull(1) ? null : r.GetString(1),
                    Author = r.GetString(2),
                    Body = r.GetString(3),
                    IsResult = r.GetInt64(4) != 0,
                    CreatedAt = ParseTime(r.GetString(5)),
                    UpdatedAt = ParseTime(r.GetString(6)),
                };
                comments[c.Id] = c;
                if (c.EntryId is { } key && entries.TryGetValue(key, out var e)) e.Comments.Add(c);
                else
                {
                    // An entry removed or re-dated in WorkHub leaves its comments on the memo.
                    c.EntryId = null;
                    memo.Comments.Add(c);
                }
            }
        }
        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT Id, CommentId, FileName, ContentType, Size, Author, CreatedAt
                  FROM WebCommentFiles WHERE MemoId = @id AND IsDeleted = 0 ORDER BY Id;
                """;
            cmd.Parameters.AddWithValue("@id", id);
            using SqliteDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (!comments.TryGetValue(r.GetInt64(1), out var c)) continue;
                var a = new TeamMemoAttachment
                {
                    Id = "c-" + r.GetInt64(0).ToString(CultureInfo.InvariantCulture),
                    MemoId = id,
                    EntryId = c.EntryId,
                    CommentId = c.Id,
                    FileName = r.GetString(2),
                    ContentType = r.GetString(3),
                    Size = r.GetInt64(4),
                    Author = r.GetString(5),
                    CreatedAt = ParseTime(r.GetString(6)),
                };
                c.Attachments.Add(a);
            }
        }
        return memo;
    }

    // ── Memos: write ───────────────────────────────────────────────────────────

    public string CreateMemo(string? folderId, string title, string kind, string user)
    {
        title = string.IsNullOrWhiteSpace(title) ? "Untitled" : title.Trim();
        string id = NewId();
        var now = DateTimeOffset.Now;
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        string path = "";
        if (folderId is not null)
        {
            using var c = Cmd(conn, tx, "SELECT Path FROM MemoCategories WHERE Id = @id;", ("@id", folderId));
            path = c.ExecuteScalar() as string ?? throw new InvalidOperationException("That folder no longer exists.");
        }
        Exec(conn, tx, """
            INSERT INTO Memos (Id, Kind, Title, Category, Content, Status, Tags, CreatedAt, UpdatedAt, StartDate,
                               IsPinned, Revision, ScheduleAllDay, RepeatWeekdays, CategoryId)
            VALUES (@id, @k, @t, @cat, '', 0, '', @now, @now, @sd, 0, 1, 0, '', @cid);
            INSERT INTO MemoBlocks (Id, MemoId, Position, Kind, Text, PayloadJson, CreatedAt, UpdatedAt, Revision)
            VALUES (@bid, @id, 0, 0, '', '{}', @now, @now, 0);
            INSERT INTO WebMemoMeta (MemoId, Owner, CreatedBy, UpdatedBy) VALUES (@id, @u, @u, @u);
            """,
            ("@id", id), ("@k", TeamMemoKind.ToWorkHub(kind)), ("@t", title), ("@cat", path), ("@now", Bin(now)),
            ("@sd", Bin(new DateTimeOffset(DateTime.Today))), ("@cid", folderId), ("@bid", NewId()), ("@u", user));
        Reindex(conn, tx, id);
        tx.Commit();
        RaiseChanged(id);
        return id;
    }

    /// <summary>
    /// Saves the header of a memo. <see cref="TeamMemo.Revision"/> must be the revision it was loaded at;
    /// a save made meanwhile (in WorkHub or by another user) raises <see cref="TeamMemoConflictException"/>.
    /// </summary>
    public void UpdateMemo(TeamMemo memo, string user)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        string path = "";
        if (memo.FolderId is not null)
        {
            using var c = Cmd(conn, tx, "SELECT Path FROM MemoCategories WHERE Id = @id;", ("@id", memo.FolderId));
            path = c.ExecuteScalar() as string ?? throw new InvalidOperationException("That folder no longer exists.");
        }
        int status = TeamMemoStatus.ToWorkHub(memo.Status);
        using (var u = Cmd(conn, tx, """
            UPDATE Memos SET Title = @t, Kind = @k, Status = @st, CategoryId = @cid, Category = @cat,
                   StartDate = @sd, DueDate = @dd, IsPinned = @pin, RepeatWeekdays = @rw, RepeatUntil = @ru,
                   PinnedAt = CASE WHEN @pin = 1 THEN IFNULL(PinnedAt, @now) ELSE NULL END,
                   CompletedAt = CASE WHEN @st = 2 THEN IFNULL(CompletedAt, @now) ELSE NULL END,
                   UpdatedAt = @now, Revision = Revision + 1
             WHERE Id = @id AND Revision = @rev;
            """,
            ("@t", string.IsNullOrWhiteSpace(memo.Title) ? "Untitled" : memo.Title.Trim()), ("@k", TeamMemoKind.ToWorkHub(memo.Kind)),
            ("@st", status), ("@cid", memo.FolderId), ("@cat", path), ("@sd", BinDate(memo.StartDate)), ("@dd", BinDate(memo.DueDate)),
            ("@pin", memo.IsPinned ? 1 : 0), ("@now", Bin(DateTimeOffset.Now)), ("@id", memo.Id), ("@rev", memo.Revision),
            ("@rw", NormalizeWeekdays(memo.RepeatWeekdays)), ("@ru", BinDate(memo.RepeatUntil))))
        {
            if (u.ExecuteNonQuery() == 0)
                throw new TeamMemoConflictException("This memo was changed (in WorkHub or by someone else) after you opened it. Reload to see the latest version, then save again.");
        }
        ReplaceTags(conn, tx, memo.Id, memo.Tags);
        UpsertMeta(conn, tx, memo.Id, user, owner: memo.Owner.Trim(), description: memo.Description);
        Reindex(conn, tx, memo.Id);
        tx.Commit();
        RaiseChanged(memo.Id);
    }

    public void SetMemoStatus(string memoId, string status, string user) =>
        TouchMemo(memoId, user, "Status = @v, CompletedAt = CASE WHEN @v = 2 THEN IFNULL(CompletedAt, @now) ELSE NULL END",
            TeamMemoStatus.ToWorkHub(status));

    public void SetPinned(string memoId, bool pinned, string user) =>
        TouchMemo(memoId, user, "IsPinned = @v, PinnedAt = CASE WHEN @v = 1 THEN @now ELSE NULL END", pinned ? 1 : 0);

    public void MoveMemo(string memoId, string? folderId, string user)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        string path = "";
        if (folderId is not null)
        {
            using var c = Cmd(conn, tx, "SELECT Path FROM MemoCategories WHERE Id = @id;", ("@id", folderId));
            path = c.ExecuteScalar() as string ?? throw new InvalidOperationException("That folder no longer exists.");
        }
        Exec(conn, tx, "UPDATE Memos SET CategoryId = @cid, Category = @cat, UpdatedAt = @now, Revision = Revision + 1 WHERE Id = @id;",
            ("@cid", folderId), ("@cat", path), ("@now", Bin(DateTimeOffset.Now)), ("@id", memoId));
        UpsertMeta(conn, tx, memoId, user);
        Reindex(conn, tx, memoId);
        tx.Commit();
        RaiseChanged(memoId);
    }

    /// <summary>Deletes the memo the way WorkHub does: the row (cascading to blocks, tags, attachments) and its files.</summary>
    public void DeleteMemo(string memoId, string user)
    {
        using SqliteConnection conn = Open();
        var files = new List<string>();
        using (var c = conn.CreateCommand())
        {
            c.CommandText = "SELECT RelativePath FROM Attachments WHERE MemoId = @id;";
            c.Parameters.AddWithValue("@id", memoId);
            using var r = c.ExecuteReader();
            while (r.Read()) files.Add(r.GetString(0));
        }
        using (SqliteTransaction tx = conn.BeginTransaction())
        {
            Exec(conn, tx, """
                DELETE FROM ReviewFacts WHERE MemoId = @id;
                DELETE FROM Memos WHERE Id = @id;
                DELETE FROM WebMemoMeta WHERE MemoId = @id;
                DELETE FROM WebEntryMeta WHERE MemoId = @id;
                UPDATE WebMemoComments SET IsDeleted = 1 WHERE MemoId = @id;
                UPDATE WebCommentFiles SET IsDeleted = 1 WHERE MemoId = @id;
                """, ("@id", memoId));
            if (HasSearchIndex(conn, tx)) Exec(conn, tx, "DELETE FROM MemoSearch WHERE MemoId = @id;", ("@id", memoId));
            tx.Commit();
        }
        foreach (string rel in files)
        {
            try { string abs = Path.Combine(AttachmentsRoot, rel); if (File.Exists(abs)) File.Delete(abs); } catch { }
        }
        RaiseChanged(memoId);
    }

    private void TouchMemo(string memoId, string user, string setClause, object value)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        Exec(conn, tx, $"UPDATE Memos SET {setClause}, UpdatedAt = @now, Revision = Revision + 1 WHERE Id = @id;",
            ("@v", value), ("@now", Bin(DateTimeOffset.Now)), ("@id", memoId));
        UpsertMeta(conn, tx, memoId, user);
        tx.Commit();
        RaiseChanged(memoId);
    }

    // ── Entries (blocks inside Memo.Content) ───────────────────────────────────

    public string AddEntry(string memoId, DateOnly date, string title, string body, DateOnly? target, string user,
        IReadOnlyList<(string Name, string ContentType, byte[] Data)>? files = null)
    {
        string key = "";
        ModifyContent(memoId, user, (conn, tx, doc) =>
        {
            DateTime when = date.ToDateTime(TimeOnly.FromDateTime(DateTime.Now));
            var block = ContentBlock.New(when, title.Trim(), (body ?? "").Trim(), target);
            doc.Blocks.Add(block);
            doc.AssignKeys();
            key = block.Key;
            foreach (var f in files ?? []) block.Images.Add((AddWorkHubAttachment(conn, tx, memoId, f.Name, f.ContentType, f.Data), ImageWidth));
            UpsertEntryMeta(conn, tx, memoId, key, user, setAuthor: true);
        });
        return key;
    }

    public void UpdateEntry(TeamMemoEntry entry, string user,
        IReadOnlyList<(string Name, string ContentType, byte[] Data)>? newFiles = null)
    {
        ModifyContent(entry.MemoId, user, (conn, tx, doc) =>
        {
            ContentBlock b = doc.Find(entry.Id) ?? throw new InvalidOperationException("That entry was removed (maybe in WorkHub). Reload the memo.");
            string oldKey = b.Key;
            DateTime when = b.IsLead || entry.EntryDate == DateOnly.FromDateTime(b.Timestamp)
                ? b.Timestamp
                : entry.EntryDate.ToDateTime(TimeOnly.FromDateTime(b.Timestamp));
            b.Edit(when, entry.Title.Trim(), (entry.Body ?? "").Trim(), entry.TargetDate, entry.DoneDate is not null);
            foreach (var f in newFiles ?? []) b.Images.Add((AddWorkHubAttachment(conn, tx, entry.MemoId, f.Name, f.ContentType, f.Data), ImageWidth));
            doc.AssignKeys();
            if (b.Key != oldKey) RekeyEntry(conn, tx, entry.MemoId, oldKey, b.Key);
            UpsertEntryMeta(conn, tx, entry.MemoId, b.Key, user, setAuthor: false);
        });
    }

    /// <summary>
    /// Marks an entry done or open. WorkHub keeps only a done flag and shows the target date as the done
    /// date, so a done entry without a target gets <paramref name="doneDate"/> as its target.
    /// </summary>
    public void SetEntryDone(string memoId, string entryKey, DateOnly? doneDate, string user)
    {
        ModifyContent(memoId, user, (conn, tx, doc) =>
        {
            ContentBlock b = doc.Find(entryKey) ?? throw new InvalidOperationException("That entry was removed (maybe in WorkHub). Reload the memo.");
            b.Edit(b.Timestamp, b.Title, b.Body, doneDate is not null ? b.Target ?? doneDate : b.Target, doneDate is not null);
            UpsertEntryMeta(conn, tx, memoId, b.Key, user, setAuthor: false);
        });
    }

    public void DeleteEntry(string memoId, string entryKey, string user)
    {
        ModifyContent(memoId, user, (conn, tx, doc) =>
        {
            ContentBlock b = doc.Find(entryKey) ?? throw new InvalidOperationException("That entry was already removed.");
            doc.Remove(b);
            Exec(conn, tx, "UPDATE WebMemoComments SET IsDeleted = 1 WHERE MemoId = @m AND EntryKey = @k;", ("@m", memoId), ("@k", entryKey));
            Exec(conn, tx, "DELETE FROM WebEntryMeta WHERE MemoId = @m AND EntryKey = @k;", ("@m", memoId), ("@k", entryKey));
            doc.AssignKeys();
        });
    }

    /// <summary>
    /// Read-modify-write of the memo body inside one immediate transaction, so a block added here never
    /// overwrites a block WorkHub saved a moment earlier.
    /// </summary>
    private void ModifyContent(string memoId, string user, Action<SqliteConnection, SqliteTransaction, ContentDoc> change)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction(deferred: false);
        string content;
        using (var c = Cmd(conn, tx, "SELECT Content FROM Memos WHERE Id = @id;", ("@id", memoId)))
            content = c.ExecuteScalar() as string ?? throw new InvalidOperationException("That memo no longer exists.");
        ContentDoc doc = ContentDoc.Parse(content);
        change(conn, tx, doc);
        Exec(conn, tx, "UPDATE Memos SET Content = @c, UpdatedAt = @now, Revision = Revision + 1 WHERE Id = @id;",
            ("@c", doc.Serialize()), ("@now", Bin(DateTimeOffset.Now)), ("@id", memoId));
        UpsertMeta(conn, tx, memoId, user);
        Reindex(conn, tx, memoId);
        tx.Commit();
        RaiseChanged(memoId);
    }

    private static void RekeyEntry(SqliteConnection conn, SqliteTransaction tx, string memoId, string oldKey, string newKey)
    {
        Exec(conn, tx, "DELETE FROM WebEntryMeta WHERE MemoId = @m AND EntryKey = @n;", ("@m", memoId), ("@n", newKey));
        Exec(conn, tx, "UPDATE WebEntryMeta SET EntryKey = @n WHERE MemoId = @m AND EntryKey = @o;", ("@m", memoId), ("@n", newKey), ("@o", oldKey));
        Exec(conn, tx, "UPDATE WebMemoComments SET EntryKey = @n WHERE MemoId = @m AND EntryKey = @o;", ("@m", memoId), ("@n", newKey), ("@o", oldKey));
    }

    // ── Comments ───────────────────────────────────────────────────────────────

    public long AddComment(string memoId, string? entryKey, string body, bool isResult, string user,
        IReadOnlyList<(string Name, string ContentType, byte[] Data)>? files = null)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        string now = Now();
        long id;
        using (var c = Cmd(conn, tx, """
            INSERT INTO WebMemoComments (MemoId, EntryKey, Author, Body, IsResult, CreatedAt, UpdatedAt)
            VALUES (@m, @e, @u, @b, @r, @now, @now); SELECT last_insert_rowid();
            """, ("@m", memoId), ("@e", entryKey), ("@u", user), ("@b", body ?? ""), ("@r", isResult ? 1 : 0), ("@now", now)))
            id = (long)c.ExecuteScalar()!;
        foreach (var f in files ?? [])
        {
            if (f.Data.LongLength > MaxAttachmentBytes) throw new InvalidOperationException($"{f.Name} is larger than 20 MB.");
            using var c = Cmd(conn, tx, """
                INSERT INTO WebCommentFiles (MemoId, CommentId, FileName, ContentType, Size, Data, Author, CreatedAt)
                VALUES (@m, @c, @n, @ct, @s, @d, @u, @now);
                """, ("@m", memoId), ("@c", id), ("@n", Path.GetFileName(f.Name)),
                ("@ct", string.IsNullOrWhiteSpace(f.ContentType) ? "application/octet-stream" : f.ContentType),
                ("@s", f.Data.LongLength), ("@d", f.Data), ("@u", user), ("@now", now));
            c.Parameters["@d"].SqliteType = SqliteType.Blob;
            c.ExecuteNonQuery();
        }
        tx.Commit();
        RaiseChanged(memoId);
        return id;
    }

    public void UpdateComment(string memoId, long commentId, string body)
    {
        Execute("UPDATE WebMemoComments SET Body = @b, UpdatedAt = @now WHERE Id = @id AND MemoId = @m;",
            ("@b", body ?? ""), ("@now", Now()), ("@id", commentId), ("@m", memoId));
        RaiseChanged(memoId);
    }

    public void SetCommentResult(string memoId, long commentId, bool isResult)
    {
        Execute("UPDATE WebMemoComments SET IsResult = @r, UpdatedAt = @now WHERE Id = @id AND MemoId = @m;",
            ("@r", isResult ? 1 : 0), ("@now", Now()), ("@id", commentId), ("@m", memoId));
        RaiseChanged(memoId);
    }

    public void DeleteComment(string memoId, long commentId)
    {
        Execute("UPDATE WebMemoComments SET IsDeleted = 1, UpdatedAt = @now WHERE Id = @id AND MemoId = @m;",
            ("@now", Now()), ("@id", commentId), ("@m", memoId));
        RaiseChanged(memoId);
    }

    // ── Attachments ────────────────────────────────────────────────────────────

    public (string FileName, string ContentType, byte[] Data)? GetAttachment(string id)
    {
        using SqliteConnection conn = Open();
        if (id.StartsWith("c-", StringComparison.Ordinal) && long.TryParse(id[2..], out long cid))
        {
            using var c = conn.CreateCommand();
            c.CommandText = "SELECT FileName, ContentType, Data FROM WebCommentFiles WHERE Id = @id AND IsDeleted = 0;";
            c.Parameters.AddWithValue("@id", cid);
            using var r = c.ExecuteReader();
            return r.Read() ? (r.GetString(0), r.GetString(1), (byte[])r.GetValue(2)) : null;
        }
        if (!id.StartsWith("w-", StringComparison.Ordinal)) return null;
        string wid = id[2..];
        string name, type, rel;
        using (var c = conn.CreateCommand())
        {
            c.CommandText = "SELECT FileName, ContentType, RelativePath FROM Attachments WHERE Id = @id;";
            c.Parameters.AddWithValue("@id", wid);
            using var r = c.ExecuteReader();
            if (!r.Read()) return null;
            (name, type, rel) = (r.GetString(0), r.GetString(1), r.GetString(2));
        }
        using (var c = conn.CreateCommand())
        {
            c.CommandText = "SELECT Content FROM AttachmentBlobs WHERE AttachmentId = @id;";
            c.Parameters.AddWithValue("@id", wid);
            if (c.ExecuteScalar() is byte[] blob && blob.Length > 0) return (name, type, blob);
        }
        string abs = Path.Combine(AttachmentsRoot, rel);
        if (!File.Exists(abs)) return null;
        using var fs = new FileStream(abs, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var ms = new MemoryStream();
        fs.CopyTo(ms);
        return (name, type, ms.ToArray());
    }

    /// <summary>
    /// Removes a comment file, or takes a picture off its history entry (the WorkHub attachment row and
    /// file stay, as WorkHub itself leaves them when a picture is taken out of the body).
    /// </summary>
    public void DeleteAttachment(string memoId, string id, string user)
    {
        if (id.StartsWith("c-", StringComparison.Ordinal) && long.TryParse(id[2..], out long cid))
        {
            Execute("UPDATE WebCommentFiles SET IsDeleted = 1 WHERE Id = @id AND MemoId = @m;", ("@id", cid), ("@m", memoId));
            RaiseChanged(memoId);
            return;
        }
        string wid = id.StartsWith("w-", StringComparison.Ordinal) ? id[2..] : id;
        ModifyContent(memoId, user, (_, _, doc) =>
        {
            foreach (ContentBlock b in doc.AllBlocks)
                if (b.Images.RemoveAll(i => i.Id == wid) > 0) b.Dirty = true;
        });
    }

    /// <summary>Stores a picture/file the WorkHub way: Attachments row + AttachmentBlobs bytes + file copy.</summary>
    private string AddWorkHubAttachment(SqliteConnection conn, SqliteTransaction tx, string memoId, string fileName, string contentType, byte[] data)
    {
        if (data.LongLength > MaxAttachmentBytes) throw new InvalidOperationException($"{fileName} is larger than 20 MB.");
        string id = NewId();
        string safe = Path.GetFileName(fileName);
        foreach (char ch in Path.GetInvalidFileNameChars()) safe = safe.Replace(ch, '_');
        if (string.IsNullOrWhiteSpace(safe)) safe = "file";
        if (safe.Length > 100) safe = safe[^100..];
        string rel = Path.Combine(memoId, Guid.NewGuid().ToString("N")[..8] + "_" + safe);
        string type = string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType;
        bool image = type.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        long now = Bin(DateTimeOffset.Now);

        Exec(conn, tx, """
            INSERT INTO Attachments (Id, MemoId, Kind, FileName, RelativePath, ContentType, SizeBytes, IsTableSource, CreatedAt, SourcePath)
            VALUES (@id, @m, @k, @n, @rel, @ct, @s, 0, @now, '');
            """, ("@id", id), ("@m", memoId), ("@k", image ? 1 : 0), ("@n", Path.GetFileName(fileName)), ("@rel", rel),
            ("@ct", type), ("@s", data.LongLength), ("@now", now));
        using (var c = Cmd(conn, tx, """
            INSERT INTO AttachmentBlobs (AttachmentId, Content, Sha256, ByteLength, StoredAt) VALUES (@id, @d, @h, @n, @now);
            """, ("@id", id), ("@d", data), ("@h", Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant()),
            ("@n", data.LongLength), ("@now", now)))
        {
            c.Parameters["@d"].SqliteType = SqliteType.Blob;
            c.ExecuteNonQuery();
        }

        // WorkHub draws pictures from the file under its attachments root.
        string abs = Path.Combine(AttachmentsRoot, rel);
        Directory.CreateDirectory(Path.GetDirectoryName(abs)!);
        File.WriteAllBytes(abs, data);
        return id;
    }

    // ── AI summaries (web side) ────────────────────────────────────────────────

    public void SaveAiSummary(string memoId, string summary, string model, string user)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        int rev;
        using (var c = Cmd(conn, tx, "SELECT Revision FROM Memos WHERE Id = @id;", ("@id", memoId)))
            rev = Convert.ToInt32(c.ExecuteScalar() ?? 0);
        Exec(conn, tx, "INSERT OR IGNORE INTO WebMemoMeta (MemoId) VALUES (@id);", ("@id", memoId));
        Exec(conn, tx, """
            UPDATE WebMemoMeta SET AiSummary = @s, AiSummaryAt = @now, AiSummaryBy = @u, AiSummaryModel = @m, AiSummaryRevision = @rev
             WHERE MemoId = @id;
            """, ("@s", summary), ("@now", Now()), ("@u", user), ("@m", model), ("@rev", rev), ("@id", memoId));
        tx.Commit();
        RaiseChanged(memoId);
    }

    public void SaveEntryAiSummary(string memoId, string entryKey, string summary, string model, string user)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        string content;
        using (var c = Cmd(conn, tx, "SELECT Content FROM Memos WHERE Id = @id;", ("@id", memoId)))
            content = c.ExecuteScalar() as string ?? "";
        string? hash = ContentDoc.Parse(content).Find(entryKey)?.Hash();
        Exec(conn, tx, "INSERT OR IGNORE INTO WebEntryMeta (MemoId, EntryKey) VALUES (@m, @k);", ("@m", memoId), ("@k", entryKey));
        Exec(conn, tx, """
            UPDATE WebEntryMeta SET AiSummary = @s, AiSummaryAt = @now, AiSummaryBy = @u, AiSummaryModel = @mo, AiSummaryHash = @h
             WHERE MemoId = @m AND EntryKey = @k;
            """, ("@s", summary), ("@now", Now()), ("@u", user), ("@mo", model), ("@h", hash), ("@m", memoId), ("@k", entryKey));
        tx.Commit();
        RaiseChanged(memoId);
    }

    // ── WorkHub bookkeeping ────────────────────────────────────────────────────

    private static void UpsertMeta(SqliteConnection conn, SqliteTransaction tx, string memoId, string user,
        string? owner = null, string? description = null)
    {
        Exec(conn, tx, "INSERT OR IGNORE INTO WebMemoMeta (MemoId, CreatedBy) VALUES (@id, '');", ("@id", memoId));
        Exec(conn, tx, """
            UPDATE WebMemoMeta SET UpdatedBy = @u,
                   Owner = IFNULL(@o, Owner), Description = IFNULL(@d, Description)
             WHERE MemoId = @id;
            """, ("@u", user), ("@o", owner), ("@d", description), ("@id", memoId));
    }

    private static void UpsertEntryMeta(SqliteConnection conn, SqliteTransaction tx, string memoId, string key, string user, bool setAuthor)
    {
        Exec(conn, tx, "INSERT OR IGNORE INTO WebEntryMeta (MemoId, EntryKey) VALUES (@m, @k);", ("@m", memoId), ("@k", key));
        Exec(conn, tx, $"UPDATE WebEntryMeta SET UpdatedAt = @now{(setAuthor ? ", Author = @u" : "")} WHERE MemoId = @m AND EntryKey = @k;",
            ("@now", Now()), ("@u", user), ("@m", memoId), ("@k", key));
    }

    /// <summary>WorkHub's ReplaceTags: Memos.Tags display text plus Tags/MemoTags rows keyed by NFKC-lowercase.</summary>
    private static void ReplaceTags(SqliteConnection conn, SqliteTransaction tx, string memoId, string csv)
    {
        var displayByNorm = (csv ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(x => x.Length > 0)
            .GroupBy(x => x.Normalize(NormalizationForm.FormKC).Trim().ToLowerInvariant(), StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
        Exec(conn, tx, "DELETE FROM MemoTags WHERE MemoId = @m;", ("@m", memoId));
        foreach (var (norm, display) in displayByNorm)
        {
            string tagId;
            using (var c = Cmd(conn, tx, "SELECT Id FROM Tags WHERE NameNorm = @n;", ("@n", norm)))
                tagId = c.ExecuteScalar() as string ?? "";
            if (tagId.Length == 0)
            {
                tagId = NewId();
                Exec(conn, tx, "INSERT INTO Tags (Id, Name, NameNorm) VALUES (@id, @name, @n);", ("@id", tagId), ("@name", display), ("@n", norm));
            }
            Exec(conn, tx, "INSERT INTO MemoTags (MemoId, TagId) VALUES (@m, @t);", ("@m", memoId), ("@t", tagId));
        }
        Exec(conn, tx, "UPDATE Memos SET Tags = @t WHERE Id = @m;", ("@t", string.Join(", ", displayByNorm.Values)), ("@m", memoId));
    }

    private static bool HasSearchIndex(SqliteConnection conn, SqliteTransaction tx)
    {
        using var c = Cmd(conn, tx, "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = 'MemoSearch';");
        return Convert.ToInt64(c.ExecuteScalar()) == 1 && TrigramCompatible(conn, tx);
    }

    private static bool? _trigramCompatible;

    /// <summary>
    /// The MemoSearch FTS5 index uses the trigram tokenizer. Writing it from an older SQLite (seen with
    /// 3.41) leaves an index WorkHub's newer SQLite reports as malformed, and WorkHub then refuses to
    /// open the DB. Below 3.45 the web leaves the index alone; WorkHub re-indexes on its next save.
    /// </summary>
    private static bool TrigramCompatible(SqliteConnection conn, SqliteTransaction tx)
    {
        if (_trigramCompatible is { } known) return known;
        using var c = Cmd(conn, tx, "SELECT sqlite_version();");
        bool ok = Version.TryParse(c.ExecuteScalar() as string, out var v) && v >= new Version(3, 45);
        _trigramCompatible = ok;
        return ok;
    }

    /// <summary>WorkHub's MemoSearchIndexer.ReindexMemo, verbatim.</summary>
    private static void Reindex(SqliteConnection conn, SqliteTransaction tx, string memoId)
    {
        if (!HasSearchIndex(conn, tx)) return;
        Exec(conn, tx, """
            DELETE FROM MemoSearch WHERE MemoId = @id;
            INSERT INTO MemoSearch (MemoId, Title, Content, Category, TagsText)
            SELECT m.Id, m.Title,
                   m.Content || ' ' || COALESCE(m.HistoryTrigger, '') || ' ' || COALESCE(m.HistoryAction, '')
                             || ' ' || COALESCE(m.HistoryResult, '') || ' ' || COALESCE(m.HistoryOwner, ''),
                   m.Category,
                   COALESCE((SELECT group_concat(t.Name, ' ') FROM MemoTags mt JOIN Tags t ON t.Id = mt.TagId WHERE mt.MemoId = m.Id), m.Tags)
            FROM Memos m WHERE m.Id = @id;
            """, ("@id", memoId));
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private void Execute(string sql, params (string Name, object? Value)[] args)
    {
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        foreach (var (name, value) in args) cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    private static SqliteCommand Cmd(SqliteConnection conn, SqliteTransaction tx, string sql, params (string Name, object? Value)[] args)
    {
        SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = sql;
        foreach (var (name, value) in args) cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        return cmd;
    }

    private static void Exec(SqliteConnection conn, SqliteTransaction tx, string sql, params (string Name, object? Value)[] args)
    {
        using SqliteCommand cmd = Cmd(conn, tx, sql, args);
        cmd.ExecuteNonQuery();
    }

    internal static string NewId() => Guid.NewGuid().ToString("N");

    private static string NormalizeWeekdays(string? csv) =>
        string.Join(",", TeamMemoCalendar.ParseWeekdays(csv).Select(d => ((int)d).ToString(CultureInfo.InvariantCulture)));

    private static string Now() => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    private static DateTime ParseTime(string s) =>
        DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var t) ? t : DateTime.MinValue;

    /// <summary>EF Core DateTimeOffsetToBinaryConverter (WorkHub's storage for every DateTimeOffset).</summary>
    internal static long Bin(DateTimeOffset v) =>
        ((v.Ticks / 1000) << 11) | ((long)v.Offset.TotalMinutes & 0x7FF);

    private static object BinDate(DateOnly? d) =>
        d is { } v ? Bin(new DateTimeOffset(v.ToDateTime(TimeOnly.MinValue))) : DBNull.Value;

    internal static DateTime FromBin(long value)
    {
        long ticks = (value >> 11) * 1000;
        int offset = (int)(value & 0x7FF);
        if (offset >= 0x400) offset -= 0x800;
        return new DateTimeOffset(ticks, TimeSpan.FromMinutes(offset)).LocalDateTime;
    }
}

public enum TeamMemoCalendarKind { Repeat, Scheduled, Due, Span }

/// <summary>One thing on one calendar day.</summary>
public sealed record TeamMemoCalendarItem(
    DateOnly Day, string MemoId, string Title, TeamMemoCalendarKind Kind, TimeOnly? Time, string Status, bool IsPlan)
{
    /// <summary>Span items: first and last day of the start → due bar.</summary>
    public DateOnly SpanStart { get; init; }
    public DateOnly SpanEnd { get; init; }
}

public static class TeamMemoCalendar
{
    private static readonly Regex LeadingTime = new(@"^\s*(\d{1,2})[:시](\d{2})?", RegexOptions.Compiled);

    public static List<DayOfWeek> ParseWeekdays(string? csv) =>
        (csv ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => int.TryParse(x, out int d) && d is >= 0 and <= 6 ? (DayOfWeek?)d : null)
            .Where(d => d is not null).Select(d => d!.Value).Distinct().OrderBy(d => (int)d).ToList();

    /// <summary>
    /// Calendar items between <paramref name="from"/> and <paramref name="to"/>, the way WorkHub's
    /// ScheduleService builds them: weekly repeats from (schedule start ?? start date ?? created) up to
    /// RepeatUntil, one-off schedules over their span, plus due dates of other memos.
    /// </summary>
    public static List<TeamMemoCalendarItem> Build(IEnumerable<TeamMemoSummary> memos, DateOnly from, DateOnly to,
        Func<TeamMemoSummary, bool> isPlan, bool includeDue)
    {
        var items = new List<TeamMemoCalendarItem>();
        foreach (TeamMemoSummary m in memos)
        {
            bool plan = isPlan(m);
            var days = ParseWeekdays(m.RepeatWeekdays);
            if (days.Count > 0)
            {
                DateOnly anchor = DateOnly.FromDateTime((m.ScheduleStartAt ?? m.StartDate?.ToDateTime(TimeOnly.MinValue) ?? m.CreatedAt).Date);
                DateOnly start = anchor > from ? anchor : from;
                DateOnly end = m.RepeatUntil is { } until && until < to ? until : to;
                TimeOnly? time = TimeOf(m);
                for (DateOnly d = start; d <= end; d = d.AddDays(1))
                    if (days.Contains(d.DayOfWeek))
                        items.Add(new(d, m.Id, m.Title, TeamMemoCalendarKind.Repeat, time, m.Status, plan));
                continue;
            }
            if (m.ScheduleStartAt is { } ss && m.ScheduleEndAt is { } se)
            {
                DateOnly a = DateOnly.FromDateTime(ss), b = DateOnly.FromDateTime(se);
                for (DateOnly d = a < from ? from : a; d <= (b > to ? to : b); d = d.AddDays(1))
                    items.Add(new(d, m.Id, m.Title, TeamMemoCalendarKind.Scheduled, m.ScheduleAllDay ? null : TimeOnly.FromDateTime(ss), m.Status, plan));
                continue;
            }
            if (includeDue && m.StartDate is { } sd && m.DueDate is { } dd && sd < dd)
            {
                // Start → due as one bar across the days.
                for (DateOnly d = sd < from ? from : sd; d <= dd && d <= to; d = d.AddDays(1))
                    items.Add(new(d, m.Id, m.Title, TeamMemoCalendarKind.Span, null, m.Status, plan) { SpanStart = sd, SpanEnd = dd });
                continue;
            }
            if (includeDue && m.DueDate is { } due && due >= from && due <= to)
                items.Add(new(due, m.Id, m.Title, TeamMemoCalendarKind.Due, null, m.Status, plan));
        }
        // Bars first, longest-running first, so one bar keeps the same row across its days.
        return items.OrderBy(i => i.Day).ThenBy(i => i.Kind == TeamMemoCalendarKind.Span ? 0 : 1)
            .ThenBy(i => i.SpanStart).ThenBy(i => i.Kind == TeamMemoCalendarKind.Span ? i.MemoId : "", StringComparer.Ordinal)
            .ThenBy(i => i.Time ?? TimeOnly.MaxValue).ThenBy(i => i.Title, StringComparer.CurrentCulture).ToList();
    }

    /// <summary>Meeting titles carry their time up front ("13:30 LPA 회의").</summary>
    private static TimeOnly? TimeOf(TeamMemoSummary m)
    {
        if (m.ScheduleStartAt is { } ss && !m.ScheduleAllDay) return TimeOnly.FromDateTime(ss);
        Match t = LeadingTime.Match(m.Title);
        if (!t.Success || !int.TryParse(t.Groups[1].Value, out int h) || h > 23) return null;
        int min = t.Groups[2].Success && int.TryParse(t.Groups[2].Value, out int mm) && mm < 60 ? mm : 0;
        return new TimeOnly(h, min);
    }
}

/// <summary>
/// Memo.Content as WorkHub writes it: optional lead text, then blocks starting with
/// <c>[yyyy-MM-dd HH:mm|target|done|dN|c]</c>. The first text line of a block is its title, the rest its
/// text; <c>!image:id:width</c> lines are pictures. Blocks the web does not touch are written back
/// byte-for-byte; only edited or new blocks are rebuilt.
/// </summary>
internal sealed class ContentDoc
{
    // Same marker as WorkHub's MemoContentFormat.EntryMarker.
    internal static readonly Regex EntryMarker = new(
        @"(?m)^\[(?<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2})(?:\|(?<target>\d{4}-\d{2}-\d{2})?)?(?:\|(?<done>done))?(?:\|d(?<depth>\d+))?(?:\|(?<collapsed>c))?\][ \t]*\r?\n?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public string Lead = "";
    /// <summary>Text before the first marker (or a body with no markers), shown as one undated entry.</summary>
    public ContentBlock? LeadBlock;
    public List<ContentBlock> Blocks = new();

    public const string LeadKey = "lead";

    /// <summary>Lead entry (if any) followed by the dated blocks.</summary>
    public IEnumerable<ContentBlock> AllBlocks => LeadBlock is null ? Blocks : Blocks.Prepend(LeadBlock);

    public static ContentDoc Parse(string content)
    {
        content ??= "";
        var doc = new ContentDoc();
        MatchCollection matches = EntryMarker.Matches(content);
        doc.Lead = matches.Count == 0 ? content : content[..matches[0].Index];
        if (doc.Lead.Trim().Length > 0) doc.LeadBlock = ContentBlock.FromLead(doc.Lead);
        for (int i = 0; i < matches.Count; i++)
        {
            Match m = matches[i];
            int end = i + 1 < matches.Count ? matches[i + 1].Index : content.Length;
            doc.Blocks.Add(ContentBlock.FromMatch(m, content[m.Index..end], content[(m.Index + m.Length)..end]));
        }
        doc.AssignKeys();
        return doc;
    }

    public ContentBlock? Find(string key) => key == LeadKey ? LeadBlock : Blocks.FirstOrDefault(b => b.Key == key);

    public void Remove(ContentBlock b)
    {
        if (ReferenceEquals(b, LeadBlock)) { LeadBlock = null; Lead = ""; }
        else Blocks.Remove(b);
    }

    /// <summary>Key = timestamp, with <c>#n</c> for the n-th repeat of the same timestamp.</summary>
    public void AssignKeys()
    {
        var seen = new Dictionary<string, int>(StringComparer.Ordinal);
        if (LeadBlock is not null) LeadBlock.Key = LeadKey;
        foreach (ContentBlock b in Blocks)
        {
            int n = seen.TryGetValue(b.Ts, out int k) ? k + 1 : 0;
            seen[b.Ts] = n;
            b.Key = n == 0 ? b.Ts : $"{b.Ts}#{n}";
        }
    }

    public string Serialize()
    {
        var sb = new StringBuilder(LeadBlock is { Dirty: true } lead ? lead.BuildLead() : Lead);
        bool prevDirty = LeadBlock?.Dirty == true;
        foreach (ContentBlock b in Blocks)
        {
            if (sb.Length > 0 && sb[^1] != '\n') sb.Append('\n');
            // A rebuilt block has no trailing blank line; keep the blank line WorkHub puts between blocks.
            if ((b.Dirty || prevDirty) && sb.Length > 0 && !EndsWithBlankLine(sb)) sb.Append('\n');
            sb.Append(b.Dirty ? b.Build() : b.Raw);
            prevDirty = b.Dirty;
        }
        return sb.ToString();
    }

    /// <summary>The text ends with an empty line; WorkHub's own blocks often use CRLF ("\r\n\r\n").</summary>
    private static bool EndsWithBlankLine(StringBuilder sb)
    {
        int i = sb.Length - 1;
        if (i < 0 || sb[i] != '\n') return false;
        i--;
        if (i >= 0 && sb[i] == '\r') i--;
        return i >= 0 && sb[i] == '\n';
    }
}

internal sealed class ContentBlock
{
    public string Raw = "";
    public bool Dirty;
    public string Ts = "";
    public DateTime Timestamp;
    public DateOnly? Target;
    public bool Done;
    public int Depth;
    public bool Collapsed;
    public string Title = "";
    public string Body = "";
    public List<(string Id, int Width)> Images = new();
    public string Key = "";

    public bool IsLead;

    public static ContentBlock FromLead(string lead)
    {
        var b = new ContentBlock { Raw = lead, IsLead = true, Ts = "", Key = ContentDoc.LeadKey };
        b.ParseInner(lead);
        return b;
    }

    /// <summary>The undated lead: title, text and pictures without a header.</summary>
    public string BuildLead()
    {
        var sb = new StringBuilder(Title);
        if (Body.Length > 0) sb.Append('\n').Append(Body);
        foreach (var (id, width) in Images) sb.Append("\n!image:").Append(id).Append(':').Append(width.ToString(CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    public static ContentBlock FromMatch(Match m, string raw, string inner)
    {
        var b = new ContentBlock
        {
            Raw = raw,
            Ts = m.Groups["timestamp"].Value,
            Done = m.Groups["done"].Success,
            Collapsed = m.Groups["collapsed"].Success,
            Depth = m.Groups["depth"].Success && int.TryParse(m.Groups["depth"].Value, out int d) ? d : 0,
        };
        b.Timestamp = DateTime.TryParseExact(b.Ts, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var t) ? t : DateTime.MinValue;
        if (m.Groups["target"].Success && DateOnly.TryParseExact(m.Groups["target"].Value, "yyyy-MM-dd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out var tg))
            b.Target = tg;

        b.ParseInner(inner);
        return b;
    }

    private void ParseInner(string inner)
    {
        var b = this;
        var lines = new List<string>();
        foreach (string line in inner.Replace("\r\n", "\n").Split('\n'))
        {
            string trimmed = line.Trim();
            if (trimmed.StartsWith("!image:", StringComparison.Ordinal))
            {
                string[] parts = trimmed["!image:".Length..].Split(':');
                if (parts[0].Length > 0)
                    b.Images.Add((parts[0], parts.Length > 1 && int.TryParse(parts[1], out int w) ? w : 440));
            }
            else lines.Add(line.TrimEnd());
        }
        // The lead may start with blank lines; its title is the first line with text.
        if (b.IsLead) while (lines.Count > 0 && lines[0].Trim().Length == 0) lines.RemoveAt(0);
        b.Title = lines.Count > 0 ? lines[0].Trim() : "";
        b.Body = string.Join("\n", lines.Skip(1)).Trim('\n');
    }

    public static ContentBlock New(DateTime when, string title, string body, DateOnly? target) => new()
    {
        Dirty = true,
        Ts = when.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
        Timestamp = when,
        Title = title,
        Body = body,
        Target = target,
    };

    public void Edit(DateTime when, string title, string body, DateOnly? target, bool done)
    {
        if (IsLead)
        {
            if (done || target is not null)
                throw new InvalidOperationException("This text has no date line in WorkHub, so it cannot carry a target or be marked done.");
            Title = title;
            Body = body;
            Dirty = true;
            return;
        }
        Ts = when.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        Timestamp = when;
        Title = title;
        Body = body;
        Target = target;
        Done = done;
        Dirty = true;
    }

    /// <summary>Header + title + text + pictures, in WorkHub's own notation (no trailing newline).</summary>
    public string Build()
    {
        var sb = new StringBuilder("[").Append(Ts);
        if (Target is { } t) sb.Append('|').Append(t.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        else if (Done) sb.Append('|');
        if (Done) sb.Append("|done");
        if (Depth > 0) sb.Append("|d").Append(Depth.ToString(CultureInfo.InvariantCulture));
        if (Collapsed) sb.Append("|c");
        sb.Append("]\n").Append(Title);
        if (Body.Length > 0) sb.Append('\n').Append(Body);
        foreach (var (id, width) in Images) sb.Append("\n!image:").Append(id).Append(':').Append(width.ToString(CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    /// <summary>Fingerprint of what an AI summary of this entry read.</summary>
    public string Hash() => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
        $"{Ts}\u0001{Target}\u0001{Done}\u0001{Title}\u0001{Body}\u0001{string.Join(",", Images.Select(i => i.Id))}")))[..16];
}
