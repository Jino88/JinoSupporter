// Unit checks for the PLATFORM_API -> BMES row-model conversions in Services/Bmes.
// Fixtures are real Tool answers (trimmed, people's names replaced) saved on 2026-10-29.
// Run: dotnet run --project PortalGNLAB.Web.BmesApiChecks   (exit code 1 when a check fails)
using PortalGNLAB.Web.Services;

var failures = new List<string>();
int checks = 0;
void Check(bool condition, string name)
{
    checks++;
    Console.WriteLine($"{(condition ? "PASS" : "FAIL")} {name}");
    if (!condition) failures.Add(name);
}

string Fixture(string name) => File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
BmesApiResult Load(string name) => BmesToolRows.ParseExecuteResponse(Fixture(name));

// -- MES020210 NG records (AI_TOOL_00055 view S) ------------------------------
{
    var result = Load("ng_view_s.json");
    var rows = BmesGateway.NgRecordRowsFromResult(result);
    Check(result.Result == "S", "NG: answer parsed as success");
    Check(rows.Count == 4, "NG: 4 fixture rows -> 4 internal rows");
    var r = rows[0];
    Check(r["WORKORDER"] == "000014861760" && r["PLANT"] == "3200", "NG: AUFNR/WERKS -> WORKORDER/PLANT");
    Check(r["PRODUCT_DATE"] == "2026-10-27", "NG: WDATE -> PRODUCT_DATE");
    Check(r["PROCESSNAME"] == "FORMING" && r["NGNAME"] == "Damage", "NG: KTSCH_TX/ZCODE_TX -> PROCESSNAME/NGNAME");
    Check(r["QTYINPUT"] == "12000" && r["QTYNG"] == "0", "NG: INQTY_O/NGQTY_O -> QTYINPUT/QTYNG");
    Check(r["PRODUCTION_LINE"] == "VP4A" && r["Shift"] == "D", "NG: VERID_TX/ZSHIF -> PRODUCTION_LINE/Shift");
    Check(!r.ContainsKey("VERID"), "NG: removed VERID column dropped");
    Check(r["USE"] == "Y" && r["LASTREGDATE"] == "2026-10-27 17:14", "NG: USEYN/ERDAT -> USE/LASTREGDATE");

    // rows with USEYN other than Y (the screen returned 'H' rows too) are kept
    var node = System.Text.Json.Nodes.JsonNode.Parse(Fixture("ng_view_s.json"))!;
    var set0 = node["resultSets"]![0]!;
    int useCol = set0["columns"]!.AsArray().Select(c => c!.GetValue<string>()).ToList().IndexOf("USEYN");
    set0["rows"]![0]![useCol] = "H";
    var kept = BmesGateway.NgRecordRowsFromResult(BmesToolRows.ParseExecuteResponse(node.ToJsonString()));
    Check(kept.Count == 4 && kept[0]["USE"] == "H", "NG: USEYN=H row kept like the screen call");

    var p = BmesGateway.NgRecordParameters("3200", new DateTime(2026, 10, 1), new DateTime(2026, 10, 27));
    Check((string)p["ZTYPE"] == "S" && (string)p["FRDAT"] == "2026-10-01" && (string)p["STDAT"] == "2026-10-27" && (string)p["WERKS"] == "3200",
        "NG: request parameters (view S, FRDAT..STDAT, WERKS)");
}

// -- MES072400 F-Cost (AI_TOOL_00001 view P) ----------------------------------
{
    var response = BmesGateway.FCostFromResult(Load("fcost_view_p.json"));
    Check(response.Rows.Count == 6, "F-Cost: 6 fixture grid rows -> 6 FCostRow");
    var inamt = response.Rows[0];
    Check(inamt.ZType == "INAMT" && inamt.Col0001 == 104672 && inamt.GetCol(14) == 3613226, "F-Cost: INAMT row COL0001..COL0014");
    Check(response.Rows[1].ZType == "FCOST" && response.Rows[1].Col0001 == 1199, "F-Cost: FCOST row value");
    Check(response.Rows[2].ZType == "FRATE" && Math.Abs(response.Rows[2].Col0001 - 1.15) < 1e-9, "F-Cost: FRATE row value");
    Check(response.Rows[1].FaccoTx == "GN" && response.Rows[1].Facco == "GN", "F-Cost: metadata forward-filled onto FCOST row");
    Check(response.Columns.Count == 14, "F-Cost: 14 period columns");
    Check(response.Columns[0].Index == 1 && response.Columns[0].Code == "20261027" && response.Columns[0].Header == "10-27",
        "F-Cost: COL0001 = 2026-10-27 (ZPVTT / ZPVTT_TX)");
    Check(response.Columns[7].Code == "2026-44" && response.Columns[11].Code == "202610", "F-Cost: week and month column codes");
    Check(response.Columns[0].Kind == FCostPeriodKind.Day && response.Columns[7].Kind == FCostPeriodKind.Week && response.Columns[11].Kind == FCostPeriodKind.Month,
        "F-Cost: day / week / month kinds inferred");
    var p = BmesGateway.FCostParameters(new DateTime(2026, 10, 27));
    Check((string)p["ZTYPE"] == "P" && (string)p["SDATE"] == "2026-10-27" && (string)p["ZGUBN"] == "D" && (string)p["FACCO"] == "GN",
        "F-Cost: request parameters (view P, daily, GN)");
}

// -- MES072410 Core-parts F-Cost (AI_TOOL_00001 view G) -----------------------
{
    var parsed = BmesGateway.CorePartsFromResult(Load("coreparts_view_g.json"));
    Check(parsed.Rows.Count == 6, "Core parts: 6 fixture grid rows");
    Check(parsed.Rows[0].Ztype == "INAMT" && parsed.Rows[0].CptypTx == "Total" && Math.Abs(parsed.Rows[0].Col0001!.Value - 31295.41) < 1e-6,
        "Core parts: INAMT Total row, COL0001");
    Check(parsed.Rows[1].Ztype == "FCOST" && parsed.Rows[1].Facco == "GN", "Core parts: FCOST row keeps/gets facility");
    Check(parsed.Columns.Count == 14 && parsed.ColumnCount == 14, "Core parts: 14 period columns");
    Check(parsed.Columns[0].Code == "20261027" && parsed.Columns[0].Kind == "Day", "Core parts: COL0001 = 2026-10-27 day");
    var p = BmesGateway.CorePartsParameters(new DateTime(2026, 10, 27));
    Check((string)p["ZTYPE"] == "G" && (string)p["ZGUBN"] == "D", "Core parts: request parameters (view G, daily)");
}

// -- MES073260 / MES073261 LPA (AI_TOOL_00027 views C and D) ------------------
{
    var (columns, rows) = BmesGateway.LpaRowsFromResult(Load("lpa_view_c.json"), "LQRNO", "AUDAT");
    Check(rows.Count == 3 && columns[0] == "LQRNO" && columns.Contains("OKCNT"), "LPA list: 3 rows, columns in Tool order");
    Check(rows[0]["LQRNO"] == "LPA-261001-001" && rows[0]["AUDAT"] == "2026-10-01" && rows[0]["TOTAL"] == "15", "LPA list: values as text");
    Check(rows[0]["MATNR"] == "" && rows[0]["lqrno"] == "LPA-261001-001", "LPA list: null -> empty, case-insensitive lookup");

    var (dColumns, dRows) = BmesGateway.LpaRowsFromResult(Load("lpa_view_d.json"), "LQRNO", "LCITM");
    Check(dRows.Count == 2 && dColumns.Contains("ZIMAG_TX"), "LPA checklist: 2 rows, photo path column present");
    Check(dRows[0]["RESUT"] == "A" && dRows[0]["LOBSQ"] == "21", "LPA checklist: RESUT / LOBSQ");

    var query = new BmesLpaScrapeService.LpaQuery { Facco = "GN", DateFrom = "20261001", DateTo = "20261028", Chker = "Hoa" };
    var p = BmesGateway.LpaListParameters(query, new DateTime(2026, 10, 1), new DateTime(2026, 10, 28));
    Check((string)p["ZTYPE"] == "C" && (string)p["SDATE"] == "2026-10-01" && (string)p["FACCO"] == "GN" && (string)p["CHKER"] == "Hoa" && !p.ContainsKey("LQRNO"),
        "LPA list: request parameters (view C, empty conditions omitted)");
    var d = BmesGateway.LpaDetailParameters(new BmesLpaScrapeService.LpaDetailQuery { Lqrno = " LPA-261001-001 " });
    Check((string)d["ZTYPE"] == "D" && (string)d["LQRNO"] == "LPA-261001-001", "LPA checklist: request parameters (view D)");
}

// -- MES073010 routing + NG codes (AI_TOOL_00055 view R + AI_TOOL_00056) -----
{
    var source = BmesGateway.RoutingSourceFromResults(Load("routing_view_r.json"), Load("codes_00056.json"));
    Check(source.Routes.Count == 6 && source.Routes.All(r => r.Matnr == "P-S-151100800"), "Routing: 6 routed steps, L99 (IPG) row left out");
    var awf = source.Routes.Single(r => r.ProcessCode == "UM03005");
    Check(awf.Werks == "3200" && awf.ProcessName == "AWF" && awf.StageCode == "L02" && awf.StageName == "Main Assy",
        "Routing: KTSCH name and ZLRTP name from AI_TOOL_00056");
    Check(source.Routes.Single(r => r.ProcessCode == "UM06186").ProcessName == "Vision Inspection Ass'y YOKE",
        "Routing: blanks in names collapsed");
    Check(source.UnnamedRoutes == 0 && source.UnnamedNgCodes == 0, "Routing: every fixture code has a name");

    Check(source.NgCodes.Count == 5, "NG codes: 5 process x NG code pairs");
    var spl = source.NgCodes.Single(c => c.ProcessCode == "UM10003" && c.NgCode == "0001");
    Check(spl.ProcessName == "Audio Bus Test" && spl.NgName == "SPL", "NG codes: names trimmed");
    Check(source.NgCodes.Single(c => c.NgCode == "0009").NgName == "Noise", "NG codes: trailing blank trimmed");

    // the same material x process on two stages keeps the smaller stage code; a process without a name is counted
    var raw = System.Text.Json.Nodes.JsonNode.Parse(Fixture("routing_view_r.json"))!;
    var input = raw["resultSets"]![0]!;
    var cols = input["columns"]!.AsArray().Select(c => c!.GetValue<string>()).ToList();
    var copy = input["rows"]![0]!.DeepClone();
    copy[cols.IndexOf("ZLRTP")] = "L01";
    input["rows"]!.AsArray().Add(copy);
    var unknown = input["rows"]![1]!.DeepClone();
    unknown[cols.IndexOf("KTSCH")] = "XX99999";
    input["rows"]!.AsArray().Add(unknown);
    var edited = BmesGateway.RoutingSourceFromResults(BmesToolRows.ParseExecuteResponse(raw.ToJsonString()), Load("codes_00056.json"));
    Check(edited.Routes.Single(r => r.ProcessCode == "UM03005").StageCode == "L01", "Routing: two stages -> smaller stage code");
    Check(edited.Routes.Count == 6 && edited.UnnamedRoutes == 1, "Routing: process without a name counted, not added");

    var p = BmesGateway.RoutingRawParameters(new DateTime(2026, 7, 28), new DateTime(2026, 10, 27));
    Check((string)p["ZTYPE"] == "R" && (string)p["FRDAT"] == "2026-07-28" && (string)p["STDAT"] == "2026-10-27"
          && ((string[])p["FACCO"]).SequenceEqual(["GN"]), "Routing: request parameters (view R, GN, 3 months)");
    var c = BmesGateway.CodeMasterParameters(new DateTime(2026, 7, 28));
    Check(((string[])c["CODTP"]).SequenceEqual(["KTSCH", "ZCODE", "ZLRTP"]) && (string)c["SDATE"] == "2026-07-28",
        "Routing: code master parameters (KTSCH with ZCODE so ZCODE gets its process name)");
}

// -- shared helpers -----------------------------------------------------------
{
    // Boundary measured on the API: 2026-07-28..2026-10-27 passes, from 2026-07-27 it is MSG0001207.
    Check(BmesToolRows.SplitPeriod(new DateTime(2026, 7, 28), new DateTime(2026, 10, 27)).Count == 1, "periods: exactly 3 months -> 1 window");
    var windows = BmesToolRows.SplitPeriod(new DateTime(2026, 4, 27), new DateTime(2026, 10, 27));
    Check(windows.Count == 3 && windows[0] == (new DateTime(2026, 4, 27), new DateTime(2026, 7, 26))
          && windows[1] == (new DateTime(2026, 7, 27), new DateTime(2026, 10, 26)) && windows[2] == (new DateTime(2026, 10, 27), new DateTime(2026, 10, 27)),
        "periods: 6 months + 1 day -> 3 back-to-back windows of at most 3 months");
    var monthEnd = BmesToolRows.SplitPeriod(new DateTime(2026, 11, 30), new DateTime(2027, 6, 30));
    Check(monthEnd.Zip(monthEnd.Skip(1)).All(p => p.Second.Start == p.First.End.AddDays(1)) && monthEnd[^1].End == new DateTime(2027, 6, 30),
        "periods: month-end start -> no gap or overlap");
    Check(BmesToolRows.SplitPeriod(new DateTime(2026, 10, 27), new DateTime(2026, 10, 27)).Count == 1, "periods: single day -> 1 window");
    Check(BmesToolRows.ToIsoDate("20261001") == "2026-10-01" && BmesToolRows.ToIsoDate("2026-10-01") == "2026-10-01" && BmesToolRows.ToIsoDate("") == "",
        "dates: yyyyMMdd -> yyyy-MM-dd");
    var empty = BmesToolRows.ParseExecuteResponse("{\"toolName\":\"AI_TOOL_00055\",\"result\":\"E\",\"message\":\"[IVDT_M_018] no data\",\"resultSets\":[]}");
    Check(empty.IsEmpty && BmesGateway.NgRecordRowsFromResult(empty).Count == 0, "answers: 'no result' (IVDT_M_018) is an empty list, not an error");
    var badParam = BmesToolRows.ParseExecuteResponse("{\"toolName\":\"AI_TOOL_00055\",\"result\":\"B\",\"message\":\"[MSG0001118]\",\"resultSets\":[]}");
    Check(!badParam.IsEmpty, "answers: parameter error is not 'no result'");
}

Console.WriteLine($"{checks - failures.Count}/{checks} checks passed");
return failures.Count == 0 ? 0 : 1;
