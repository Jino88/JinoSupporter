using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace PortalGNLAB.Web.Services;

public sealed class RequestAiRequest
{
    public string? Text { get; set; }
    public string? Today { get; set; }
    public string? Kind { get; set; }
    public WorkingRequestSave? Current { get; set; }
    public List<PlaceAiImage>? Images { get; set; }
}

public sealed class RequestAiAnswer
{
    public WorkingRequestSave Request { get; set; } = new();
    public string Note { get; set; } = "";
}

public sealed record RequestAiResult(WorkingRequestSave Request, string Note, string Model);

/// <summary>
/// Working > Request > AI input: the user pastes a JIG / equipment list (Excel cells, a mail) or pictures of one and the
/// Codex API fills the PROCESSING REQUEST form. Nothing is saved here; the client shows the form and the user saves it.
/// </summary>
public sealed partial class WorkingRequestAiService(CodexApiService codex)
{
    public const string Effort = "medium";
    /// <summary>A picture of a JIG/MC list may be up to 20 MB (the other AI inputs keep 5 MB).</summary>
    public const int MaxImageBytes = 20 * 1024 * 1024;

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public async Task<RequestAiResult> FillAsync(RequestAiRequest body, CancellationToken ct = default)
    {
        string request = body.Text?.Trim() ?? "";
        List<PlaceAiPicture> pictures = WorkingPlaceAiService.ReadImages(body.Images, MaxImageBytes);
        if (request.Length == 0 && pictures.Count == 0) throw new ArgumentException("Type or paste what to put in.");
        if (!codex.IsConfigured) throw new ArgumentException("Codex API is not configured. Set the API key in Setting.");

        JsonArray content = BuildContent(body, pictures);

        string model = codex.Model;
        CodexTextResult result = await codex.GenerateAsync(Instructions, content, model, Effort, 16000, ct);
        codex.RecordUsage(result.Usage);

        RequestAiAnswer answer = ParseAnswer(result.Text, body);
        return new RequestAiResult(answer.Request, answer.Note, $"{model} · {Effort}");
    }

    /// <summary>The AI input: one text block (kind, current form, picture list, user text) then each picture as input_image.</summary>
    public static JsonArray BuildContent(RequestAiRequest body, IReadOnlyList<PlaceAiPicture> pictures)
    {
        WorkingRequestSave current = WorkingRequestService.Normalize(body.Current);
        if (!string.IsNullOrWhiteSpace(body.Kind) && body.Current is null) current.Kind = WorkingRequestService.Text(body.Kind, 50);
        var text = new System.Text.StringBuilder();
        text.AppendLine($"Today: {TodayOf(body)}");
        text.AppendLine();
        text.AppendLine("=== CURRENT FORM (JSON) ===");
        text.AppendLine(JsonSerializer.Serialize(current, Json));
        text.AppendLine();
        if (pictures.Count > 0)
        {
            text.AppendLine("=== PICTURES (numbered from 0, in the order they follow) ===");
            for (int i = 0; i < pictures.Count; i++) text.AppendLine($"Picture {i}: {pictures[i].Name}");
            text.AppendLine();
        }
        text.AppendLine("=== INPUT ===");
        text.AppendLine((body.Text ?? "").Trim());
        var content = new JsonArray { new JsonObject { ["type"] = "input_text", ["text"] = text.ToString() } };
        foreach (PlaceAiPicture p in pictures)
            content.Add(new JsonObject { ["type"] = "input_image", ["image_url"] = $"data:{p.MediaType};base64,{p.Base64}" });
        return content;
    }

    /// <summary>
    /// Reads the AI answer leniently: strips code fences, takes the outermost {...} ("request" object or the form fields
    /// at the top), fills what the AI left empty from the current form, then normalizes.
    /// </summary>
    public static RequestAiAnswer ParseAnswer(string? text, RequestAiRequest? request)
    {
        const string fail = "The AI answer could not be read. Try again.";
        string s = text ?? "";
        int a = s.IndexOf('{'), b = s.LastIndexOf('}');
        if (a < 0 || b <= a) throw new ArgumentException(fail);
        try
        {
            using JsonDocument doc = JsonDocument.Parse(s[a..(b + 1)]);
            JsonElement root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) throw new ArgumentException(fail);

            JsonElement form = root;
            foreach (JsonProperty p in root.EnumerateObject())
                if (string.Equals(p.Name, "request", StringComparison.OrdinalIgnoreCase) && p.Value.ValueKind == JsonValueKind.Object) form = p.Value;
            string note = Str(root, "note");

            WorkingRequestSave current = WorkingRequestService.Normalize(request?.Current);
            bool hasCurrent = request?.Current is not null;
            string kind = hasCurrent ? current.Kind : Fallback(WorkingRequestService.Text(request?.Kind, 50), current.Kind);

            var items = new List<WorkingRequestItem>();
            if (Find(form, "items") is { ValueKind: JsonValueKind.Array } arr)
                foreach (JsonElement e in arr.EnumerateArray())
                {
                    if (e.ValueKind != JsonValueKind.Object) throw new ArgumentException(fail);
                    items.Add(ReadItem(e));
                }
            if (items.Count == 0) items = current.Items;

            var save = new WorkingRequestSave
            {
                Kind = Fallback(Str(form, "kind"), kind),
                Requester = Fallback(Str(form, "requester"), current.Requester),
                Phone = Fallback(Str(form, "phone"), current.Phone),
                Checker = Fallback(Str(form, "checker"), current.Checker),
                Confirm = Fallback(Str(form, "confirm"), current.Confirm),
                Method = Fallback(Str(form, "method"), current.Method),
                DepartmentName = Fallback(Str(form, "departmentName"), current.DepartmentName),
                EquipmentName = Fallback(Str(form, "equipmentName"), current.EquipmentName),
                PartName = Fallback(Str(form, "partName"), current.PartName),
                ItemApplication = Fallback(Str(form, "itemApplication"), current.ItemApplication),
                UsingPlace = Fallback(Str(form, "usingPlace"), current.UsingPlace),
                RequestDate = Fallback(WorkingRequestService.DateOf(Str(form, "requestDate")), current.RequestDate),
                FinishDate1 = Fallback(WorkingRequestService.DateOf(Str(form, "finishDate1")), current.FinishDate1),
                FinishDate2 = Fallback(WorkingRequestService.DateOf(Str(form, "finishDate2")), current.FinishDate2),
                Purpose = Fallback(Str(form, "purpose"), current.Purpose),
                Others = Fallback(Str(form, "others"), current.Others),
                // the code is given by the server when the request is saved; the AI never makes one up
                DeptCode = current.DeptCode,
                Qty1 = Fallback(Str(form, "qty1"), current.Qty1),
                Qty2 = Fallback(Str(form, "qty2"), current.Qty2),
                Items = items,
            };
            save = WorkingRequestService.Normalize(save);
            if (save.PartName.Length == 0) save.PartName = CommonModel(save.Items);
            return new RequestAiAnswer { Request = save, Note = (note ?? "").Trim() };
        }
        catch (JsonException) { throw new ArgumentException(fail); }
    }

    /// <summary>The model every row has in common ("" when the rows have none or different ones).</summary>
    public static string CommonModel(IEnumerable<WorkingRequestItem> items)
    {
        List<string> models = items.Select(i => i.Model.Trim()).Where(m => m.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        return models.Count == 1 ? WorkingRequestService.Text(models[0], 200) : "";
    }

    private static string Fallback(string value, string fallback) => value.Length > 0 ? value : fallback;

    private static WorkingRequestItem ReadItem(JsonElement e)
    {
        double? lineSpec = Num(e, "lineSpec"), stock = Num(e, "stock"), order = Num(e, "order");
        double line = lineSpec ?? 0, have = stock ?? 0;
        return new WorkingRequestItem
        {
            No = (int)(Num(e, "no") ?? 0),
            Model = Str(e, "model"),
            Name = Str(e, "name"),
            LineSpec = line,
            Stock = have,
            // no order given: what the line needs minus what is in stock, never below 0
            Order = order ?? Math.Max(0, line - have),
            DrawingNo = Str(e, "drawingNo"),
            Picture = Str(e, "picture"),
            Note = Str(e, "note"),
        };
    }

    private static JsonElement? Find(JsonElement e, string name)
    {
        foreach (JsonProperty p in e.EnumerateObject())
            if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)) return p.Value;
        return null;
    }

    /// <summary>A property as text: strings as they are, numbers and booleans as written, anything else "".</summary>
    private static string Str(JsonElement e, string name) => Find(e, name) switch
    {
        { ValueKind: JsonValueKind.String } v => (v.GetString() ?? "").Trim(),
        { ValueKind: JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False } v => v.GetRawText(),
        _ => "",
    };

    /// <summary>A property as a number: numbers, or text such as "12", "1,200" or "3 ea"; null when missing or not a number.</summary>
    private static double? Num(JsonElement e, string name)
    {
        switch (Find(e, name))
        {
            case { ValueKind: JsonValueKind.Number } v: return v.TryGetDouble(out double d) ? d : null;
            case { ValueKind: JsonValueKind.String } v:
                Match m = LeadingNumber().Match((v.GetString() ?? "").Replace(",", "").Trim());
                return m.Success && double.TryParse(m.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double n) ? n : null;
            default: return null;
        }
    }

    [GeneratedRegex(@"^-?\d+(\.\d+)?")]
    private static partial Regex LeadingNumber();

    private static string TodayOf(RequestAiRequest? request)
    {
        string d = WorkingRequestService.DateOf(request?.Today);
        return d.Length > 0 ? d : DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    internal const string Instructions = """
        You fill in a JIG / MC PROCESSING REQUEST form (a request to the machining shop for JIGs or equipment). The input is
        today's date, the current form as JSON and the user's text. The text may be a mail, cells copied from Excel (tab
        separated) or pictures of a JIG / equipment list. Read it and answer the whole form.

        The form (JSON, camelCase):
        - kind: "JIG/MC" unless the input says otherwise
        - requester, phone, checker, confirm: people and phone number as written
        - method: "new" (make new), "repair" or "scrap"; default "new"
        - departmentName: default "ME GROUP"
        - equipmentName: the equipment or line the items are for; partName: the part / model name
        - itemApplication, usingPlace: what the items are used for and where
        - requestDate, finishDate1, finishDate2: yyyy-MM-dd (a date without a year takes the year of today), or ""
        - purpose: default "New request"; others: remarks
        - qty1, qty2: quantity texts as written, or ""
        - items: one row per JIG / equipment: {"no":1,"model":"","name":"","lineSpec":0,"stock":0,"order":0,"drawingNo":"","picture":"","note":""}
          - model: the product model the item belongs to; name: the JIG / equipment name (required)
          - lineSpec: quantity the line needs; stock: quantity in stock; order: quantity to order
          - if the order quantity is not given, order = lineSpec - stock (0 when that is below 0)
          - drawingNo: drawing number (도번); picture: picture reference as written; note: remark of the row
          - numbers only (decimals allowed, never negative)

        Rules:
        - Answer ONLY a JSON object: {"request":{...the whole form...},"note":"..."} with no other text.
        - When the current form has values, keep them and only fill in or change what the input tells you. Answer all items of the form (the current ones plus new ones; do not repeat a row that is already there).
        - Do not invent values that are not in the input; use "" (or 0 for numbers). Leave deptCode "".
        - If partName is not given and all items share one model, partName is that model.
        - note: 1 to 3 short English sentences saying what you filled in and anything left out or unclear.
        """;
}
