namespace PortalGNLAB.Web.Services;

/// <summary>
/// Model Info › Documents: what an uploaded workbook is, so several files picked at once each go to their place —
/// a JIG / Machine / Tester list (its JIG LIST, Machine List or Tester List sheets), Work Instructions (a sheet per
/// process), or neither.
/// </summary>
public static class WorkingDocumentKind
{
    public const string JigList = "jig-list", Instructions = "instructions", Unknown = "unknown";

    public static string Of(Stream xlsx)
    {
        var data = new MemoryStream();
        xlsx.CopyTo(data);
        try
        {
            data.Position = 0;
            if (WorkingProcessListService.ReadItems(data).Count > 0) return JigList;
        }
        catch (ArgumentException) { }

        string tmp = Path.Combine(Path.GetTempPath(), $"doc-kind-{Guid.NewGuid():N}.xlsx");
        try
        {
            File.WriteAllBytes(tmp, data.ToArray());
            return WorkingInstructionService.ParseFile(tmp).Processes.Count > 0 ? Instructions : Unknown;
        }
        catch (ArgumentException) { return Unknown; }
        finally { try { File.Delete(tmp); } catch { } }
    }
}
