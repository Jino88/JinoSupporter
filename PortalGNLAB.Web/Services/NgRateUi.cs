namespace PortalGNLAB.Web.Services;

/// <summary>Display helpers shared by the NG Rate report tables.</summary>
public static class NgRateUi
{
    /// <summary>CSS class that colours a process type badge (Sub / Main / Function / Visual / OQC / other).</summary>
    public static string ProcessTypeClass(string? processType)
    {
        string t = (processType ?? "").Trim().ToUpperInvariant();
        if (t.StartsWith("SUB")) return "pt-sub";
        if (t.StartsWith("MAIN")) return "pt-main";
        if (t.StartsWith("FUNC")) return "pt-func";
        if (t.Contains("VISUAL") || t.StartsWith("VIS")) return "pt-vis";
        if (t.StartsWith("OQC")) return "pt-oqc";
        return "pt-other";
    }
}
