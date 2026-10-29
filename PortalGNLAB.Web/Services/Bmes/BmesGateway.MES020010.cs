using System.Text.Json;

namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    /// <summary>
    /// MES020010/SearchList — material master.
    /// Input: logged-in session (all materials, SAVYN=Y, full date range).
    /// Output: every material row, or null when the call failed or returned no rows; the
    /// reason is reported through <paramref name="progress"/> as an "[ERROR] ..." line.
    /// </summary>
    public static async Task<List<BmesMaterial>?> FetchMaterialsAsync(
        BmesWebSession session, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        if (await WebLoginErrorAsync(session, progress, cancellationToken) is { } loginError)
        {
            progress?.Report($"[ERROR] {loginError}");
            return null;
        }

        progress?.Report("Calling MES020010/SearchList…");
        string url =
            $"{BmesWebSession.BaseUrl}/MES020010/SearchList?perPage=" +
            "&Condition.MATNR=&Condition.MAKTX=&Condition.SAVYN=Y&Condition.KEYWO=" +
            "&Condition.GRCOD=&Condition.GRCOD_TX=&Condition.MTYPE=&Condition.BTYPE=&Condition.INJTP=" +
            "&Condition.SDATE=1900-01-01&Condition.EDATE=2050-12-31&page=1";

        HttpResponseMessage resp;
        try   { resp = await session.Client.GetAsync(url, cancellationToken); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Request failed: {ex.Message}");
            return null;
        }
        if (!resp.IsSuccessStatusCode)
        {
            progress?.Report($"[ERROR] HTTP {(int)resp.StatusCode} {resp.StatusCode}");
            return null;
        }

        string json = await resp.Content.ReadAsStringAsync(cancellationToken);
        try
        {
            using var doc = JsonDocument.Parse(json);
            JsonElement contents = doc.RootElement.GetProperty("data").GetProperty("contents");
            if (contents.ValueKind != JsonValueKind.Array || contents.GetArrayLength() == 0)
            {
                progress?.Report("[ERROR] No contents in response.");
                return null;
            }

            string fetchedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var list = new List<BmesMaterial>(contents.GetArrayLength());
            foreach (var item in contents.EnumerateArray())
            {
                list.Add(new BmesMaterial
                {
                    Matnr     = ReadRowText(item, "MATNR"),
                    Maktx     = ReadRowText(item, "MAKTX"),
                    Meins     = ReadRowText(item, "MEINS"),
                    Injtp     = ReadRowText(item, "INJTP"),
                    Mtype     = ReadRowText(item, "MTYPE"),
                    Btype     = ReadRowText(item, "BTYPE"),
                    MngCode   = ReadRowText(item, "MNGCODE").Trim(),
                    ModNameB  = ReadRowText(item, "MODNAME_B"),
                    LotQt     = ReadRowText(item, "LOTQT"),
                    Bunch     = ReadRowText(item, "BUNCH"),
                    NgTar     = ReadRowText(item, "NGTAR"),
                    McLv1Tx   = ReadRowText(item, "MCLV1_TX"),
                    McLv2Tx   = ReadRowText(item, "MCLV2_TX"),
                    McLv3Tx   = ReadRowText(item, "MCLV3_TX"),
                    McLv4Tx   = ReadRowText(item, "MCLV4_TX"),
                    McLv5Tx   = ReadRowText(item, "MCLV5_TX"),
                    McLv6Tx   = ReadRowText(item, "MCLV6_TX"),
                    Ernam     = ReadRowText(item, "ERNAM"),
                    Erdat     = ReadRowText(item, "ERDAT"),
                    Grcod     = ReadRowText(item, "GRCOD"),
                    Grnam     = ReadRowText(item, "GRNAM"),
                    MfPhi     = ReadRowText(item, "MFPHI"),
                    FetchedAt = fetchedAt,
                });
            }
            return list;
        }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Parse failed: {ex.Message}");
            return null;
        }
    }

    // Shared by the screen partials in this folder owned by the same refactor step.
    private static string ReadRowText(JsonElement obj, string key)
    {
        if (!obj.TryGetProperty(key, out var prop)) return string.Empty;
        return prop.ValueKind switch
        {
            JsonValueKind.Null   => string.Empty,
            JsonValueKind.String => prop.GetString() ?? string.Empty,
            _                    => prop.ToString(),
        };
    }
}
