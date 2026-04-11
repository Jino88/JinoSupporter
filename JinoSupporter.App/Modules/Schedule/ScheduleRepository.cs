using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using WorkbenchHost.Infrastructure;

namespace JinoSupporter.App.Modules.Schedule;

public sealed class ScheduleItem
{
    public long      Id          { get; init; }
    public string    Title       { get; init; } = string.Empty;
    public string    Description { get; init; } = string.Empty;
    public DateOnly  StartDate   { get; init; }
    public DateOnly  EndDate     { get; init; }
    public TimeOnly? StartTime   { get; init; }
    public TimeOnly? EndTime     { get; init; }
    public string    Color       { get; init; } = "#4A90D9";
    public List<string> Tags     { get; init; } = [];
    public string    CreatedAt   { get; init; } = string.Empty;

    public bool IsMultiDay => EndDate > StartDate;
    public bool IsAllDay   => StartTime is null;

    public string TimeDisplay => IsAllDay ? "종일"
        : EndTime.HasValue
            ? $"{StartTime!.Value.ToString("HH:mm")}~{EndTime.Value.ToString("HH:mm")}"
            : StartTime!.Value.ToString("HH:mm");

    public string TagsDisplay => Tags.Count > 0 ? string.Join(", ", Tags) : string.Empty;
}

public sealed class ScheduleRepository
{
    public static string DatabasePath
    {
        get
        {
            string custom = WorkbenchSettingsStore.GetScheduleDatabasePath();
            return string.IsNullOrWhiteSpace(custom)
                ? Path.Combine(WorkbenchSettingsStore.DefaultStorageRootDirectory, "schedule.db")
                : custom;
        }
    }

    public static string DefaultDatabasePath =>
        Path.Combine(WorkbenchSettingsStore.DefaultStorageRootDirectory, "schedule.db");

    private SqliteConnection OpenConnection()
    {
        var conn = new SqliteConnection($"Data Source={DatabasePath}");
        conn.Open();
        return conn;
    }

    public ScheduleRepository()
    {
        string? dir = Path.GetDirectoryName(DatabasePath);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS Schedules (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                Title       TEXT    NOT NULL DEFAULT '',
                Description TEXT    NOT NULL DEFAULT '',
                StartDate   TEXT    NOT NULL,
                EndDate     TEXT    NOT NULL,
                Color       TEXT    NOT NULL DEFAULT '#4A90D9',
                Tags        TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL,
                StartTime   TEXT    NULL,
                EndTime     TEXT    NULL
            );
            CREATE INDEX IF NOT EXISTS idx_sched_start ON Schedules(StartDate);
            CREATE INDEX IF NOT EXISTS idx_sched_end   ON Schedules(EndDate);
            """;
        cmd.ExecuteNonQuery();
        MigrateSchema(conn);
    }

    private static void MigrateSchema(SqliteConnection conn)
    {
        using SqliteCommand check = conn.CreateCommand();
        check.CommandText = "PRAGMA table_info(Schedules);";
        bool hasTags = false, hasStartTime = false, hasEndTime = false;
        using (SqliteDataReader r = check.ExecuteReader())
        {
            while (r.Read())
            {
                string col = r.GetString(1);
                if (col.Equals("Tags",      StringComparison.OrdinalIgnoreCase)) hasTags      = true;
                if (col.Equals("StartTime", StringComparison.OrdinalIgnoreCase)) hasStartTime = true;
                if (col.Equals("EndTime",   StringComparison.OrdinalIgnoreCase)) hasEndTime   = true;
            }
        }
        if (!hasTags)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE Schedules ADD COLUMN Tags TEXT NOT NULL DEFAULT '[]';";
            alter.ExecuteNonQuery();
        }
        if (!hasStartTime)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE Schedules ADD COLUMN StartTime TEXT NULL;";
            alter.ExecuteNonQuery();
        }
        if (!hasEndTime)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE Schedules ADD COLUMN EndTime TEXT NULL;";
            alter.ExecuteNonQuery();
        }
    }

    // ── Write ────────────────────────────────────────────────────────────────

    public long AddSchedule(string title, string description,
                            DateOnly startDate, DateOnly endDate,
                            string color, IEnumerable<string>? tags = null,
                            TimeOnly? startTime = null, TimeOnly? endTime = null)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO Schedules (Title, Description, StartDate, EndDate, Color, Tags, CreatedAt, StartTime, EndTime)
            VALUES (@title, @desc, @start, @end, @color, @tags, @at, @startTime, @endTime);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@title",     title);
        cmd.Parameters.AddWithValue("@desc",      description);
        cmd.Parameters.AddWithValue("@start",     startDate.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@end",       endDate.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@color",     color);
        cmd.Parameters.AddWithValue("@tags",      JsonSerializer.Serialize(tags?.ToList() ?? []));
        cmd.Parameters.AddWithValue("@at",        DateTime.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("@startTime", startTime.HasValue ? (object)startTime.Value.ToString("HH:mm") : DBNull.Value);
        cmd.Parameters.AddWithValue("@endTime",   endTime.HasValue   ? (object)endTime.Value.ToString("HH:mm")   : DBNull.Value);
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public void UpdateSchedule(long id, string title, string description,
                               DateOnly startDate, DateOnly endDate,
                               string color, IEnumerable<string>? tags = null,
                               TimeOnly? startTime = null, TimeOnly? endTime = null)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE Schedules
            SET Title=@title, Description=@desc, StartDate=@start,
                EndDate=@end, Color=@color, Tags=@tags,
                StartTime=@startTime, EndTime=@endTime
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@id",        id);
        cmd.Parameters.AddWithValue("@title",     title);
        cmd.Parameters.AddWithValue("@desc",      description);
        cmd.Parameters.AddWithValue("@start",     startDate.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@end",       endDate.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@color",     color);
        cmd.Parameters.AddWithValue("@tags",      JsonSerializer.Serialize(tags?.ToList() ?? []));
        cmd.Parameters.AddWithValue("@startTime", startTime.HasValue ? (object)startTime.Value.ToString("HH:mm") : DBNull.Value);
        cmd.Parameters.AddWithValue("@endTime",   endTime.HasValue   ? (object)endTime.Value.ToString("HH:mm")   : DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    public void DeleteSchedule(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Schedules WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    // ── Read ─────────────────────────────────────────────────────────────────

    /// <summary>startDate ~ endDate 범위와 겹치는 스케쥴 조회.</summary>
    public List<ScheduleItem> GetSchedulesInRange(DateOnly from, DateOnly to,
                                                  IReadOnlyList<string>? filterTags = null)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Title, Description, StartDate, EndDate, Color, Tags, CreatedAt, StartTime, EndTime
            FROM Schedules
            WHERE StartDate <= @to AND EndDate >= @from
            ORDER BY StartDate, IFNULL(StartTime, ''), EndDate;
            """;
        cmd.Parameters.AddWithValue("@from", from.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@to",   to.ToString("yyyy-MM-dd"));

        var list = new List<ScheduleItem>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            List<string> tags = [];
            try { tags = JsonSerializer.Deserialize<List<string>>(r.GetString(6)) ?? []; } catch { }

            TimeOnly? startTime = null, endTime = null;
            if (!r.IsDBNull(8) && TimeOnly.TryParse(r.GetString(8), out TimeOnly st)) startTime = st;
            if (!r.IsDBNull(9) && TimeOnly.TryParse(r.GetString(9), out TimeOnly et)) endTime   = et;

            list.Add(new ScheduleItem
            {
                Id          = r.GetInt64(0),
                Title       = r.GetString(1),
                Description = r.GetString(2),
                StartDate   = DateOnly.Parse(r.GetString(3)),
                EndDate     = DateOnly.Parse(r.GetString(4)),
                Color       = r.GetString(5),
                Tags        = tags,
                CreatedAt   = r.GetString(7),
                StartTime   = startTime,
                EndTime     = endTime
            });
        }

        // 태그 필터 적용
        if (filterTags is { Count: > 0 })
        {
            var tagSet = new HashSet<string>(filterTags, StringComparer.OrdinalIgnoreCase);
            list = list.Where(s => s.Tags.Any(tagSet.Contains)).ToList();
        }

        return list;
    }

    /// <summary>오늘부터 n일 이내 스케쥴.</summary>
    public List<ScheduleItem> GetUpcoming(int days = 7, IReadOnlyList<string>? filterTags = null)
    {
        DateOnly today = DateOnly.FromDateTime(DateTime.Today);
        DateOnly until = today.AddDays(days - 1);
        return GetSchedulesInRange(today, until, filterTags);
    }

    /// <summary>DB에 저장된 모든 고유 태그 목록.</summary>
    public List<string> GetAllDistinctTags()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM Schedules;";

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var tags = JsonSerializer.Deserialize<List<string>>(r.GetString(0)) ?? [];
                foreach (string t in tags)
                    if (!string.IsNullOrWhiteSpace(t)) set.Add(t.Trim());
            }
            catch { }
        }
        return [.. set.OrderBy(t => t, StringComparer.OrdinalIgnoreCase)];
    }
}
