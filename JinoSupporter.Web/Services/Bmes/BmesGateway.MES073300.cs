using System.Text;
using System.Text.Json;

namespace JinoSupporter.Web.Services;

public static partial class BmesGateway
{
    /// <summary>
    /// Search condition captured from the BMES UI. PDTNO is sent as observed — the response
    /// comes back with many different PDTNO values, so it does not appear to narrow the set.
    /// </summary>
    private const string PdmSearchPayload =
        """
        {"Condition":{"LASTY":"L","SDATE":"1900-01-01","EDATE":"2050-12-31","PDMNO":"","PITTX":"","PDRTX":"","BMOTX":"","USEYN":"Y","PDTNO":"0001"}}
        """;

    // All columns observed in the MES073300/SearchList response.
    public static readonly string[] PdmColumns =
    {
        "APSTA", "APSTA_TX",
        "PDMNO", "PDMVR", "PDMV",
        "PDTNO", "PDTNO_TX",
        "PDRTX", "PITTX", "PRRTX",
        "BMONO", "BMONO_TX", "BMOTX",
        "USEYN", "DELYN", "APPYN", "DETYN", "REVYN", "MODYN", "CWFYN",
        "APKEY", "DESCR",
        "ERNAM", "ERNAM_TX", "ERDAT",
        "ARNAM_TX", "ARDAT",
    };

    /// <summary>
    /// MES073300/SearchList — drawing/spec master (PDM).
    /// Input: logged-in session (fixed search condition, see <c>PdmSearchPayload</c>).
    /// Output: every row as column→text dictionaries (keys = <see cref="PdmColumns"/>), or null
    /// when the request or parse failed (reason reported through <paramref name="progress"/>).
    /// </summary>
    public static async Task<List<Dictionary<string, string>>?> FetchPdmRowsAsync(
        BmesWebSession session, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        progress?.Report("Calling MES073300/SearchList…");
        string json;
        try
        {
            using var content = new StringContent(PdmSearchPayload, Encoding.UTF8, "application/json");
            var resp = await session.Client.PostAsync($"{BmesWebSession.BaseUrl}/MES073300/SearchList", content, cancellationToken);
            if (!resp.IsSuccessStatusCode)
            {
                progress?.Report($"[ERROR] HTTP {(int)resp.StatusCode} {resp.StatusCode}");
                return null;
            }
            json = await resp.Content.ReadAsStringAsync(cancellationToken);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            progress?.Report($"[ERROR] Request failed: {ex.Message}");
            return null;
        }

        var rows = new List<Dictionary<string, string>>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("data", out var dataEl) ||
                !dataEl.TryGetProperty("contents", out var contents) ||
                contents.ValueKind != JsonValueKind.Array)
            {
                progress?.Report("[ERROR] Response has no data.contents array.");
                return null;
            }

            foreach (var item in contents.EnumerateArray())
            {
                var row = new Dictionary<string, string>(PdmColumns.Length);
                foreach (var col in PdmColumns)
                    row[col] = ReadRowText(item, col);
                rows.Add(row);
            }
        }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Parse failed: {ex.Message}");
            return null;
        }
        return rows;
    }

    /// <summary>
    /// MES073300 SearchDetail → DownloadCheck → Download — fetches the drawing attached to a PDMNO.
    /// Input: session opened with Referrer=/MES073300 and AjaxHeader=false (the header is added
    /// only to the two JSON calls), the trimmed non-empty PDMNO and its revision (blank = "0").
    /// Output: file name and bytes, or a failed result with the reason.
    /// </summary>
    public static async Task<BmesPdmDownload> DownloadPdmDrawingAsync(
        BmesWebSession session, string docNo, string pdmvr, IProgress<string>? progress = null)
    {
        string version = string.IsNullOrWhiteSpace(pdmvr) ? "0" : pdmvr.Trim();

        // 1. SearchDetail — locate the attached file.
        progress?.Report("Reading drawing detail…");
        string detailJson;
        try
        {
            string payload = JsonSerializer.Serialize(new
            {
                Condition = new { PDMNO = docNo, PDMVR = version },
            });
            var resp = await PostPdmJsonAsync(session.Client, $"{BmesWebSession.BaseUrl}/MES073300/SearchDetail", payload);
            if (!resp.IsSuccessStatusCode)
                return BmesPdmDownload.Failed($"SearchDetail HTTP {(int)resp.StatusCode}.");
            detailJson = await resp.Content.ReadAsStringAsync();
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("SearchDetail failed: " + ex.Message);
        }

        string pattx, pattp, folder, physical, fileVersion;
        try
        {
            using var doc = JsonDocument.Parse(detailJson);
            if (!doc.RootElement.TryGetProperty("Files", out var files) ||
                files.ValueKind != JsonValueKind.Array ||
                files.GetArrayLength() == 0)
            {
                return BmesPdmDownload.Failed($"{docNo} has no attached file.");
            }

            // Prefer a usable DWG; otherwise fall back to the first attachment.
            JsonElement pick = files[0];
            foreach (var file in files.EnumerateArray())
            {
                bool usable = !string.Equals(ReadRowText(file, "USEYN"), "N", StringComparison.OrdinalIgnoreCase);
                if (usable && ReadRowText(file, "PATTP").Equals("DWG", StringComparison.OrdinalIgnoreCase))
                {
                    pick = file;
                    break;
                }
            }

            pattx = ReadRowText(pick, "PATTX");
            pattp = ReadRowText(pick, "PATTP");
            folder = ReadRowText(pick, "P_FOLDE");
            physical = ReadRowText(pick, "P_PATTX");

            // The file carries its own revision, which is not always the version asked for.
            fileVersion = ReadRowText(pick, "PDMVR");
            if (string.IsNullOrWhiteSpace(fileVersion))
                fileVersion = version;
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("SearchDetail parse failed: " + ex.Message);
        }

        // 2. DownloadCheck — the server refuses the file unless this passes first.
        progress?.Report($"Checking download permission for {pattx}…");
        try
        {
            string payload = JsonSerializer.Serialize(new
            {
                Download = new
                {
                    PDMNO = docNo,
                    PDMVR = fileVersion,
                    PATTX = pattx,
                    P_FOLDE = folder,
                    P_PATTX = physical,
                },
            });
            var resp = await PostPdmJsonAsync(session.Client, $"{BmesWebSession.BaseUrl}/MES073300/DownloadCheck", payload);
            string body = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode)
                return BmesPdmDownload.Failed($"DownloadCheck HTTP {(int)resp.StatusCode}.");

            using var doc = JsonDocument.Parse(body);
            string result = doc.RootElement.TryGetProperty("Result", out var r) ? r.GetString() ?? "" : "";
            if (!string.Equals(result, "S", StringComparison.OrdinalIgnoreCase))
            {
                string msg = doc.RootElement.TryGetProperty("Msg", out var m) ? m.GetString() ?? "" : "";
                return BmesPdmDownload.Failed($"DownloadCheck refused: {result} {msg}".Trim());
            }
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("DownloadCheck failed: " + ex.Message);
        }

        // 3. The file itself — a form-encoded POST, not JSON, using Download[...] field names.
        progress?.Report($"Downloading {pattx}…");
        try
        {
            using var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["Download[PDMNO]"] = docNo,
                ["Download[PDMVR]"] = fileVersion,
                ["Download[PATTX]"] = pattx,
                ["Download[P_FOLDE]"] = folder,
                ["Download[P_PATTX]"] = physical,
            });
            var resp = await session.Client.PostAsync($"{BmesWebSession.BaseUrl}/MES073300/Download", form);
            if (!resp.IsSuccessStatusCode)
                return BmesPdmDownload.Failed($"Download HTTP {(int)resp.StatusCode}.");

            // A failure comes back as an HTML error page rather than a non-200 status.
            string mediaType = resp.Content.Headers.ContentType?.MediaType ?? string.Empty;
            if (resp.Content.Headers.ContentDisposition is null &&
                mediaType.Contains("html", StringComparison.OrdinalIgnoreCase))
            {
                return BmesPdmDownload.Failed("Download returned an HTML page instead of a file.");
            }

            byte[] bytes = await resp.Content.ReadAsByteArrayAsync();
            if (bytes.Length == 0)
                return BmesPdmDownload.Failed("Download returned an empty body.");

            string fileName = resp.Content.Headers.ContentDisposition?.FileNameStar
                ?? resp.Content.Headers.ContentDisposition?.FileName?.Trim('"')
                ?? (string.IsNullOrWhiteSpace(pattx) ? $"{docNo}.{pattp}" : pattx);

            return BmesPdmDownload.Ok(fileName, bytes);
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("Download failed: " + ex.Message);
        }
    }

    private static Task<HttpResponseMessage> PostPdmJsonAsync(HttpClient client, string url, string payload)
    {
        var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json"),
        };
        request.Headers.Add("X-Requested-With", "XMLHttpRequest");
        return client.SendAsync(request);
    }
}
