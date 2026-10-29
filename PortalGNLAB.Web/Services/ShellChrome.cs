using Microsoft.AspNetCore.WebUtilities;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// A page opened in its own window carries <c>chrome=off</c>, and the shell layouts then leave
/// out the navigation rail and the tab strip. The popup exists to show one record; the menu
/// inside it would only offer ways to navigate away from it, and the window has no tab bar of
/// its own to get back with.
/// </summary>
public static class ShellChrome
{
    public const string QueryKey = "chrome";
    public const string HiddenValue = "off";

    /// <summary>Ready to append to a URL that already has a query string.</summary>
    public const string HiddenQuery = "&" + QueryKey + "=" + HiddenValue;

    public static bool IsHidden(string uri)
    {
        int cut = uri.IndexOf('?');
        if (cut < 0)
            return false;

        return QueryHelpers.ParseNullableQuery(uri[cut..]) is { } query
            && query.TryGetValue(QueryKey, out var values)
            && values.Any(value => string.Equals(value, HiddenValue, StringComparison.OrdinalIgnoreCase));
    }
}
