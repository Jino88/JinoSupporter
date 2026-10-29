using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using ClosedXML.Excel;

namespace JinoSupporter.Web.Services;

/// <summary>
/// AI Summary for a team memo: sends the memo header, every history entry, every comment and the
/// attachments (images as images, spreadsheets/text as text, PDFs as files) to the Codex API and
/// stores the markdown answer on the memo. Runs Codex Luna at low reasoning effort.
/// </summary>
public sealed class TeamMemoAiSummaryService(TeamMemoService memos, CodexApiService codex, ILogger<TeamMemoAiSummaryService> log)
{
    public const string Model = CodexApiService.LunaModel;
    public const string Effort = "low";

    private const int MaxImages = 16;
    private const long ImageBudgetBytes = 2_500_000;
    private const int MaxPdfs = 3;
    private const long MaxPdfBytes = 10L * 1024 * 1024;
    private const int MaxTextPerFile = 20_000;
    private const int MaxTextTotal = 80_000;
    private const int MaxSheetRows = 200;
    private const int MaxSheetCols = 30;

    // One summary per memo / per entry at a time across all users; a second click is refused.
    // Keys are "m{memoId}" and "e{entryId}".
    private static readonly ConcurrentDictionary<string, string> Running = new();

    public bool IsConfigured => codex.IsConfigured;

    /// <summary>User name running a summary of this memo right now, or null.</summary>
    public static string? RunningBy(string memoId) => Running.TryGetValue($"m{memoId}", out var u) ? u : null;

    /// <summary>User name running a summary of this history entry right now, or null.</summary>
    public static string? RunningByEntry(string memoId, string entryKey) => Running.TryGetValue($"e{memoId}/{entryKey}", out var u) ? u : null;

    public async Task<string> SummarizeAsync(string memoId, string user, CancellationToken ct = default)
    {
        string key = $"m{memoId}";
        if (!Running.TryAdd(key, user))
            throw new InvalidOperationException($"{RunningBy(memoId)} is already running AI Summary for this memo.");
        try
        {
            TeamMemo memo = memos.GetMemo(memoId) ?? throw new InvalidOperationException("That memo no longer exists.");
            JsonArray content = BuildContent(memo, null);
            log.LogInformation("Team memo AI summary: memo={MemoId} user={User} parts={Parts}", memoId, user, content.Count);

            CodexTextResult result = await codex.GenerateAsync(Instructions, content, Model, Effort, 8000, ct);
            codex.RecordUsage(result.Usage);
            memos.SaveAiSummary(memoId, result.Text, $"{result.Model} · {Effort}", user);
            log.LogInformation("Team memo AI summary done: memo={MemoId} chars={Chars} usage={In}/{Out}",
                memoId, result.Text.Length, result.Usage?.InputTokens, result.Usage?.OutputTokens);
            return result.Text;
        }
        finally
        {
            Running.TryRemove(key, out _);
        }
    }

    /// <summary>Summarizes one history entry: its text, every comment on it and its/their attachments.</summary>
    public async Task<string> SummarizeEntryAsync(string memoId, string entryId, string user, CancellationToken ct = default)
    {
        string key = $"e{memoId}/{entryId}";
        if (!Running.TryAdd(key, user))
            throw new InvalidOperationException($"{RunningByEntry(memoId, entryId)} is already running AI Summary for this entry.");
        try
        {
            TeamMemo memo = memos.GetMemo(memoId) ?? throw new InvalidOperationException("That memo no longer exists.");
            TeamMemoEntry entry = memo.Entries.FirstOrDefault(e => e.Id == entryId)
                                  ?? throw new InvalidOperationException("That history entry no longer exists.");
            JsonArray content = BuildContent(memo, entry);
            log.LogInformation("Team memo entry AI summary: memo={MemoId} entry={EntryId} user={User} parts={Parts}",
                memoId, entryId, user, content.Count);

            CodexTextResult result = await codex.GenerateAsync(EntryInstructions, content, Model, Effort, 6000, ct);
            codex.RecordUsage(result.Usage);
            memos.SaveEntryAiSummary(memoId, entryId, result.Text, $"{result.Model} · {Effort}", user);
            log.LogInformation("Team memo entry AI summary done: entry={EntryId} chars={Chars} usage={In}/{Out}",
                entryId, result.Text.Length, result.Usage?.InputTokens, result.Usage?.OutputTokens);
            return result.Text;
        }
        finally
        {
            Running.TryRemove(key, out _);
        }
    }

    private const string EntryInstructions = """
        You summarize ONE dated history entry of a shared engineering work memo (manufacturing /
        quality / process review): the entry text, every user's comments on it, and the attached
        images, spreadsheets, text files and PDFs. The memo header is given only as context.
        Read the images carefully: they are usually screenshots of result tables, charts, defect
        photos or drawings — extract the actual numbers, NG counts/rates and conclusions they show
        instead of describing the picture.

        Write in the language most of the entry text is written in (Korean if mixed with Korean; if the
        entry has almost no text, use the language of the memo header).
        Output GitHub-flavored markdown, short, with these parts (skip a part if there is nothing for it):

        **What was done** — 1-2 sentences.
        **Data** — the key values read from the images/files. When an image is a table, reproduce the
        important rows as a compact markdown table (keep the original column names and units).
        **Result** — the conclusion; comments marked RESULT take priority.
        **Open / next** — what is still open, the target date if overdue, and who is on it.

        Rules: do not invent facts, values or owners not in the input; write [unclear] for unreadable
        parts. Keep product codes, process names, model names and units exactly as written.
        """;

    private const string Instructions = """
        You summarize a shared engineering work memo (manufacturing / quality / process review).
        The input is the memo header, its dated history entries with every user's comments, and the
        attached images, spreadsheets, text files and PDFs. Read the images carefully: they are often
        screenshots of tables, charts, inspection results or drawings — extract the numbers and
        conclusions they show instead of describing the picture.

        Write the summary in the language most of the memo text is written in (Korean if mixed with Korean).
        Output GitHub-flavored markdown with exactly these sections, skipping a section only if there is nothing for it:

        ## Overview
        2-4 sentences: what this work is, why it is done, where it stands now.
        ## Timeline
        One bullet per history entry in date order: `MM-DD` what was done → outcome (Done / Open / overdue).
        ## Key results & data
        Results marked RESULT first, then numbers and conclusions read from tables, images and files.
        Use a small markdown table when comparing values. Cite the source like (IMG 3) or (file: name.xlsx).
        ## Open issues & next steps
        Unfinished entries, overdue targets, unanswered questions, and who is on them.

        Rules: do not invent facts, values or owners that are not in the input; write [unclear] for
        unreadable parts. Keep product codes, model names and units exactly as written. Be concise.
        """;

    /// <summary>
    /// Message parts for the whole memo, or — when <paramref name="only"/> is set — for that one entry
    /// with the memo header as context.
    /// </summary>
    private JsonArray BuildContent(TeamMemo memo, TeamMemoEntry? only)
    {
        var parts = new JsonArray();
        var text = new StringBuilder();
        int imageNo = 0, pdfCount = 0, textBudget = MaxTextTotal;
        var pendingImages = new List<(int No, TeamMemoAttachment A)>();
        var pendingPdfs = new List<TeamMemoAttachment>();

        string AttachmentRefs(List<TeamMemoAttachment> list)
        {
            if (list.Count == 0) return "";
            var refs = new List<string>();
            foreach (TeamMemoAttachment a in list)
            {
                if (a.IsImage && imageNo < MaxImages)
                {
                    imageNo++;
                    pendingImages.Add((imageNo, a));
                    refs.Add($"IMG {imageNo} ({a.FileName})");
                }
                else if (IsPdf(a) && pendingPdfs.Count < MaxPdfs && a.Size <= MaxPdfBytes)
                {
                    pendingPdfs.Add(a);
                    refs.Add($"file: {a.FileName} (attached PDF)");
                }
                else
                {
                    string? body = textBudget > 0 ? ExtractText(a) : null;
                    if (body is not null)
                    {
                        if (body.Length > Math.Min(MaxTextPerFile, textBudget))
                            body = body[..Math.Min(MaxTextPerFile, textBudget)] + "\n…(truncated)";
                        textBudget -= body.Length;
                        refs.Add($"file: {a.FileName}");
                        text.AppendLine($"<file name=\"{a.FileName}\">");
                        text.AppendLine(body);
                        text.AppendLine("</file>");
                    }
                    else
                    {
                        refs.Add($"file: {a.FileName} ({a.ContentType}, not readable)");
                    }
                }
            }
            return "Attachments: " + string.Join("; ", refs);
        }

        text.AppendLine($"# Memo: {memo.Title}");
        text.AppendLine($"Type: {memo.Kind} | Status: {TeamMemoStatus.Label(memo.Status)} | Owner: {Or(memo.Owner)} | Start: {D(memo.StartDate)} | Due: {D(memo.DueDate)} | Today: {D(DateOnly.FromDateTime(DateTime.Today))}");
        if (!string.IsNullOrWhiteSpace(memo.Tags)) text.AppendLine($"Tags: {memo.Tags}");
        if (!string.IsNullOrWhiteSpace(memo.Description)) text.AppendLine($"Summary written by users:\n{memo.Description.Trim()}");
        if (only is null)
        {
            string memoRefs = AttachmentRefs(memo.Attachments);
            if (memoRefs.Length > 0) text.AppendLine(memoRefs);
        }

        text.AppendLine();
        text.AppendLine(only is null ? "## History entries" : "## The history entry to summarize");
        if (memo.Entries.Count == 0) text.AppendLine("(none)");
        foreach (TeamMemoEntry e in only is null ? memo.Entries : [only])
        {
            string state = e.IsDone ? $"Done {D(e.DoneDate)}"
                : e.TargetDate is { } t ? $"Open, target {D(t)}{(t < DateOnly.FromDateTime(DateTime.Today) ? " (overdue)" : "")}"
                : "Open";
            text.AppendLine();
            text.AppendLine($"### [{D(e.EntryDate)}] {(e.Depth > 0 ? new string('>', e.Depth) + " sub-entry: " : "")}{e.Title} — by {e.Author} — {state}");
            if (!string.IsNullOrWhiteSpace(e.Body)) text.AppendLine(e.Body.Trim());
            string refs = AttachmentRefs(e.Attachments);
            if (refs.Length > 0) text.AppendLine(refs);
            AppendComments(text, e.Comments, AttachmentRefs);
        }

        if (only is null && memo.Comments.Count > 0)
        {
            text.AppendLine();
            text.AppendLine("## Discussion on the whole memo");
            AppendComments(text, memo.Comments, AttachmentRefs);
        }

        parts.Add(new JsonObject { ["type"] = "input_text", ["text"] = text.ToString() });

        foreach ((int no, TeamMemoAttachment a) in pendingImages)
        {
            if (memos.GetAttachment(a.Id) is not { } file) continue;
            (string b64, string mediaType) = (Convert.ToBase64String(file.Data), file.ContentType);
            if (OperatingSystem.IsWindowsVersionAtLeast(6, 1) && file.Data.LongLength > ImageBudgetBytes)
            {
                try { (b64, mediaType) = ImageCompressor.CompressIfLarge(b64, mediaType, ImageBudgetBytes); }
                catch (Exception ex) { log.LogWarning(ex, "Could not shrink team memo image {Id}", a.Id); }
            }
            parts.Add(new JsonObject { ["type"] = "input_text", ["text"] = $"IMG {no}: {a.FileName}" });
            parts.Add(new JsonObject
            {
                ["type"] = "input_image",
                ["image_url"] = $"data:{mediaType};base64,{b64}",
                ["detail"] = "high",
            });
        }

        foreach (TeamMemoAttachment a in pendingPdfs)
        {
            if (pdfCount >= MaxPdfs || memos.GetAttachment(a.Id) is not { } file) continue;
            pdfCount++;
            parts.Add(new JsonObject
            {
                ["type"] = "input_file",
                ["filename"] = a.FileName,
                ["file_data"] = $"data:application/pdf;base64,{Convert.ToBase64String(file.Data)}",
            });
        }
        return parts;
    }

    private static void AppendComments(StringBuilder text, List<TeamMemoComment> comments, Func<List<TeamMemoAttachment>, string> refs)
    {
        if (comments.Count == 0) return;
        text.AppendLine("Comments:");
        foreach (TeamMemoComment c in comments)
        {
            string flag = c.IsResult ? " [RESULT]" : "";
            text.AppendLine($"- {c.Author} ({c.CreatedAt:yyyy-MM-dd HH:mm}){flag}: {c.Body.Trim().Replace("\n", "\n  ")}");
            string r = refs(c.Attachments);
            if (r.Length > 0) text.AppendLine("  " + r);
        }
    }

    private static bool IsPdf(TeamMemoAttachment a) =>
        a.ContentType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase)
        || a.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);

    private static readonly string[] TextExtensions = [".txt", ".csv", ".tsv", ".md", ".json", ".log", ".xml", ".html", ".htm"];

    /// <summary>Plain text of a spreadsheet or text attachment, or null when it is neither.</summary>
    private string? ExtractText(TeamMemoAttachment a)
    {
        string ext = Path.GetExtension(a.FileName).ToLowerInvariant();
        bool isText = a.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) || TextExtensions.Contains(ext);
        bool isExcel = ext is ".xlsx" or ".xlsm";
        if (!isText && !isExcel) return null;
        if (memos.GetAttachment(a.Id) is not { } file) return null;
        try
        {
            return isExcel ? WorkbookToText(file.Data) : Encoding.UTF8.GetString(file.Data);
        }
        catch (Exception ex)
        {
            log.LogWarning(ex, "Could not read team memo attachment {Id} as text", a.Id);
            return null;
        }
    }

    private static string WorkbookToText(byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var wb = new XLWorkbook(ms);
        var sb = new StringBuilder();
        foreach (IXLWorksheet ws in wb.Worksheets)
        {
            IXLRange? used = ws.RangeUsed();
            if (used is null) continue;
            sb.AppendLine($"[sheet {ws.Name}] (tab-separated)");
            int rows = Math.Min(used.RowCount(), MaxSheetRows);
            int cols = Math.Min(used.ColumnCount(), MaxSheetCols);
            for (int r = 1; r <= rows; r++)
            {
                var cells = new string[cols];
                for (int c = 1; c <= cols; c++)
                    cells[c - 1] = used.Cell(r, c).GetFormattedString().Replace('\t', ' ').Replace('\n', ' ');
                sb.AppendLine(string.Join('\t', cells).TrimEnd('\t'));
            }
            if (used.RowCount() > rows) sb.AppendLine($"…({used.RowCount() - rows} more rows)");
        }
        return sb.ToString();
    }

    private static string D(DateOnly? d) => d is { } v ? v.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "-";
    private static string Or(string s) => string.IsNullOrWhiteSpace(s) ? "-" : s;
}
