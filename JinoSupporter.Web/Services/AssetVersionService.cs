using System.Collections.Concurrent;
using System.Security.Cryptography;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Cache-busting token for a static asset, derived from the file's own content.
///
/// The alternative is a hand-written <c>?v=3</c> in the markup, which has to be
/// bumped by whoever edits the asset. Forgetting is silent: the server serves the
/// new file and the browser keeps the old one, so a CSS fix appears not to work
/// and the next person debugs the CSS instead of the cache. A content hash cannot
/// be forgotten.
///
/// Hashes are cached for the lifetime of the process. Editing an asset while the
/// app runs will not change its token until the next start — which is what a
/// rebuild does anyway.
/// </summary>
public sealed class AssetVersionService(IWebHostEnvironment env)
{
    private readonly ConcurrentDictionary<string, string> _tokens = new(StringComparer.Ordinal);

    /// <param name="webRootRelativePath">Path under wwwroot, e.g. <c>ui-redesign/assets/instrument.scoped.css</c>.</param>
    /// <returns>A short hex token, or <c>"0"</c> when the file cannot be read.</returns>
    public string For(string webRootRelativePath) =>
        _tokens.GetOrAdd(webRootRelativePath, Compute);

    private string Compute(string relative)
    {
        try
        {
            string root = env.WebRootPath;
            if (string.IsNullOrEmpty(root)) return "0";

            string full = Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));
            using FileStream stream = File.OpenRead(full);

            // 4 bytes is 8 hex characters — plenty to separate revisions of one
            // file, and short enough to stay readable in the rendered markup.
            return Convert.ToHexString(SHA256.HashData(stream), 0, 4).ToLowerInvariant();
        }
        catch (IOException) { return "0"; }
        catch (UnauthorizedAccessException) { return "0"; }
        catch (ArgumentException) { return "0"; }
    }
}
