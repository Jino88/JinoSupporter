using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.Versioning;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Compresses images that exceed Claude's 5 MB per-image limit.
/// Quality-first strategy: reduce JPEG quality (near-lossless → visibly lossy) before
/// touching dimensions, so resolution stays intact as long as possible.
/// </summary>
[SupportedOSPlatform("windows6.1")]
public static class ImageCompressor
{
    // Default per-image target when sending a SINGLE image: 3.5 MB of raw bytes
    // (~4.67 MB base64), safely under Anthropic's per-image cap.
    public const long DefaultTargetRawBytes = 3_500_000;

    // Anthropic vision hard cap: 8000 px on any side. Pre-scale any image whose
    // longer side exceeds this BEFORE byte-budget compression, or the request
    // is rejected with HTTP 400.
    public const int MaxDimensionPx = 7900;

    // Quality ladder: start near-lossless, step down only if needed.
    private static readonly int[] QualityLadder = [95, 90, 85, 80, 75, 65];

    /// <summary>
    /// Computes a per-image raw-bytes budget that keeps the TOTAL request under
    /// Anthropic's vision-request size limit (≈8 MB base64). Divides ~5.5 MB raw
    /// budget across the image count, clamping each image to a reasonable minimum.
    /// </summary>
    public static long BudgetPerImage(int imageCount)
    {
        if (imageCount <= 1) return DefaultTargetRawBytes;              // 3.5 MB
        // Keep total base64 under ~7.2 MB → raw ≈ 5.4 MB total
        const long TotalRawBudget = 5_400_000;
        long perImage = TotalRawBudget / imageCount;
        // Floor to 900 KB (enough for readable text at moderate resolution)
        return Math.Max(900_000, perImage);
    }

    /// <summary>
    /// If <paramref name="base64"/> decodes to more than the target,
    /// re-encodes to JPEG with the highest quality that fits. Downscales dimensions
    /// only when even the lowest quality still exceeds the target.
    /// Returns the new (base64, mediaType) — or the originals if already small enough.
    /// </summary>
    public static (string Base64, string MediaType) CompressIfLarge(
        string base64, string mediaType, long? targetRawBytes = null)
    {
        long target = targetRawBytes ?? DefaultTargetRawBytes;

        byte[] data;
        try { data = Convert.FromBase64String(base64); }
        catch { return (base64, mediaType); }

        int probeW = 0, probeH = 0;
        try
        {
            using var probeMs = new MemoryStream(data);
            using var probe   = Image.FromStream(probeMs);
            probeW = probe.Width;
            probeH = probe.Height;
        }
        catch { }

        bool dimensionTooLarge = probeW > MaxDimensionPx || probeH > MaxDimensionPx;
        if (data.Length <= target && !dimensionTooLarge) return (base64, mediaType);

        try
        {
            using var inputMs = new MemoryStream(data);
            using var src     = Image.FromStream(inputMs);

            int w = src.Width;
            int h = src.Height;

            // Pre-scale if any side exceeds Anthropic's 8000 px hard limit.
            if (w > MaxDimensionPx || h > MaxDimensionPx)
            {
                double capScale = (double)MaxDimensionPx / Math.Max(w, h);
                w = Math.Max(600, (int)(w * capScale));
                h = Math.Max(400, (int)(h * capScale));
            }

            // Try quality-only compression at full resolution first.
            foreach (int q in QualityLadder)
            {
                byte[] bytes = EncodeJpeg(src, w, h, q);
                if (bytes.Length <= target)
                    return (Convert.ToBase64String(bytes), "image/jpeg");
            }

            // If even q65 full-resolution is too big, downscale gradually at q85 (visually clean).
            double scale = 0.85;
            for (int attempt = 0; attempt < 6; attempt++)
            {
                int nw = Math.Max(600, (int)(w * scale));
                int nh = Math.Max(400, (int)(h * scale));
                byte[] bytes = EncodeJpeg(src, nw, nh, 85);
                if (bytes.Length <= target)
                    return (Convert.ToBase64String(bytes), "image/jpeg");
                scale *= 0.8;
            }

            // Last resort — aggressive: quality 70 + 50% dimensions.
            byte[] final = EncodeJpeg(src, Math.Max(600, w / 2), Math.Max(400, h / 2), 70);
            return (Convert.ToBase64String(final), "image/jpeg");
        }
        catch
        {
            // Resize failed (e.g. unsupported format on non-Windows) — return original.
            return (base64, mediaType);
        }
    }

    private static byte[] EncodeJpeg(Image src, int width, int height, int quality)
    {
        using var bmp = new Bitmap(width, height);
        using (var g = Graphics.FromImage(bmp))
        {
            g.InterpolationMode  = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode      = SmoothingMode.HighQuality;
            g.PixelOffsetMode    = PixelOffsetMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.Clear(Color.White);
            g.DrawImage(src, 0, 0, width, height);
        }

        using var outMs = new MemoryStream();
        ImageCodecInfo jpgEncoder = ImageCodecInfo.GetImageEncoders()
            .First(e => e.FormatID == ImageFormat.Jpeg.Guid);
        var encParams = new EncoderParameters(1);
        encParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);
        bmp.Save(outMs, jpgEncoder, encParams);
        return outMs.ToArray();
    }
}
