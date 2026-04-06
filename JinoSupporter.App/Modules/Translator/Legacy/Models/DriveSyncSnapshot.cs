namespace CustomKeyboardCSharp.Models;

public sealed class DriveSyncSnapshot
{
    public int SchemaVersion { get; set; } = 3;
    public long ExportedAt { get; set; }
    public List<DriveSyncHistoryRecord> Histories { get; set; } = [];
    public List<DriveSyncOptionRecord> Options { get; set; } = [];
    public List<DriveSyncVocabularyRecord> Vocabulary { get; set; } = [];
}

public sealed class DriveSyncHistoryRecord
{
    public long Id { get; set; }
    public long CreatedAt { get; set; }
    public string Provider { get; set; } = string.Empty;
    public string Mode { get; set; } = string.Empty;
    public string Direction { get; set; } = string.Empty;
    public string SourceText { get; set; } = string.Empty;
    public string ResultJson { get; set; } = string.Empty;
}

public sealed class DriveSyncOptionRecord
{
    public long Id { get; set; }
    public long HistoryId { get; set; }
    public int Position { get; set; }
    public string TranslatedText { get; set; } = string.Empty;
    public string Nuance { get; set; } = string.Empty;
}

public sealed class DriveSyncVocabularyRecord
{
    public long Id { get; set; }
    public long CreatedAt { get; set; }
    public string Provider { get; set; } = string.Empty;
    public string Mode { get; set; } = string.Empty;
    public string Direction { get; set; } = string.Empty;
    public string SourceWord { get; set; } = string.Empty;
    public string TargetMeaning { get; set; } = string.Empty;
    public string SourceText { get; set; } = string.Empty;
    public long? HistoryId { get; set; }
}
