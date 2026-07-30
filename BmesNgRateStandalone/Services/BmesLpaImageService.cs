namespace BmesNgRateStandalone.Services;

/// <summary>
/// Serves the LPA result photos (MES073261 <c>ZIMAG_TX</c>) to the viewer one at a time,
/// on demand. The originals are phone photos — 3.7 MB on average, 7.5 MB at the top end —
/// so instead of downloading every photo of a search up front (which blocked Search for
/// tens of seconds), the viewer points each <c>&lt;img&gt;</c> at <c>/bmes/lpa/img</c> and
/// this fetches, shrinks and stores each one the first time it is actually shown.
///
/// The two downscaled sizes — a table thumbnail and a popup view — are kept in the app DB
/// (<c>BmesLpaImages</c>, via <see cref="WebRepository"/>), keyed by path, so a photo is
/// downloaded from BMES exactly once: later views, later searches and restarts all read the
/// stored copy.
///
/// The endpoint needs NO session: <c>/MES073261/GetImage</c> is served to anonymous
/// requests, so this never touches <see cref="BmesLpaScrapeService"/>'s login handshake.
/// </summary>
public sealed class BmesLpaImageService(WebRepository repo)
{
    /// <summary>Public (no login) image URL; the viewer uses the same prefix for the "Original"
    /// link that opens the untouched original.</summary>
    public const string ImageUrlPrefix = "https://bmes.bujeon.com/MES073261/GetImage?fileName=";

    /// <summary>Longest side, in pixels. The thumbnail sits in a table cell; the view is
    /// what the popup shows, sized to fill a normal screen without paying for a 12 MP
    /// original nobody zooms into.</summary>
    private const int ThumbMaxPx = 160;
    private const int ViewMaxPx  = 1000;

    private const long ThumbQuality = 70L;
    private const long ViewQuality  = 72L;

    /// <summary>One client for every image request: the endpoint is anonymous, so there is
    /// no per-user cookie to isolate, and a single shared client avoids socket exhaustion
    /// when a viewer loads a screenful of thumbnails at once.</summary>
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(120) };

    /// <summary>
    /// One photo at the requested size, ready to write to the response. Returns the stored
    /// copy if there is one; otherwise downloads the original once, re-encodes BOTH sizes,
    /// stores the pair and returns the size asked for.
    ///
    /// Null on any failure (deleted photo answering with an HTML error page, network error,
    /// non-Windows host) so the caller can answer 404 and the browser falls back to the
    /// live "Original" link.
    /// </summary>
    public async Task<byte[]?> GetAsync(string path, bool view)
    {
        path = path.Trim();
        if (path.Length == 0) return null;

        byte[]? stored = repo.GetLpaImage(path, view);
        if (stored is not null) return stored;

        (byte[]? thumb, byte[]? viewBytes) = await DownloadPairAsync(path);
        if (thumb is null || viewBytes is null) return null;

        try { repo.SaveLpaImagePair(path, thumb, viewBytes); }
        catch { /* storage is an optimisation, not a requirement */ }

        return view ? viewBytes : thumb;
    }

    /// <summary>Download the original from BMES and re-encode it at both sizes.</summary>
    private static async Task<(byte[]? Thumb, byte[]? View)> DownloadPairAsync(string path)
    {
        byte[] original;
        try
        {
            using var response = await Http.GetAsync(ImageUrlPrefix + Uri.EscapeDataString(path));
            if (!response.IsSuccessStatusCode) return (null, null);
            // A deleted file comes back as an HTML error page with 200, so trust the type.
            string? type = response.Content.Headers.ContentType?.MediaType;
            if (type is null || !type.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                return (null, null);
            original = await response.Content.ReadAsByteArrayAsync();
        }
        catch { return (null, null); }

        byte[]? thumb = Resize(original, ThumbMaxPx, ThumbQuality);
        byte[]? view  = Resize(original, ViewMaxPx, ViewQuality);
        return (thumb, view);
    }

    /// <summary>
    /// Longest side down to <paramref name="maxPx"/>, re-encoded as JPEG. Never enlarges —
    /// a small original stays as it is, just re-encoded.
    /// </summary>
    private static byte[]? Resize(byte[] data, int maxPx, long quality)
    {
        // System.Drawing is Windows-only (6.1+, which is what the analyzer checks); the guard
        // makes a non-Windows host degrade to "no image" instead of crashing.
        if (!OperatingSystem.IsWindowsVersionAtLeast(6, 1)) return null;

        try
        {
            using var input = new MemoryStream(data);
            using var source = System.Drawing.Image.FromStream(input);

            double scale = Math.Min(1.0, (double)maxPx / Math.Max(source.Width, source.Height));
            int w = Math.Max(1, (int)Math.Round(source.Width * scale));
            int h = Math.Max(1, (int)Math.Round(source.Height * scale));

            using var target = new System.Drawing.Bitmap(w, h);
            using (var g = System.Drawing.Graphics.FromImage(target))
            {
                g.InterpolationMode  = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode    = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.SmoothingMode      = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                // White behind transparent PNGs: JPEG has no alpha and the default is black.
                g.Clear(System.Drawing.Color.White);
                g.DrawImage(source, 0, 0, w, h);
            }

            // Plain loop, not FirstOrDefault: the platform guard above does not flow into a
            // lambda, so LINQ here would light up CA1416 on every member it touches.
            Guid jpeg = System.Drawing.Imaging.ImageFormat.Jpeg.Guid;
            System.Drawing.Imaging.ImageCodecInfo? encoder = null;
            foreach (var codec in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
                if (codec.FormatID == jpeg) { encoder = codec; break; }
            if (encoder is null) return null;

            using var parameters = new System.Drawing.Imaging.EncoderParameters(1);
            parameters.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, quality);

            using var output = new MemoryStream();
            target.Save(output, encoder, parameters);
            return output.ToArray();
        }
        catch { return null; }
    }
}
