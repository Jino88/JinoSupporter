namespace PortalGNLAB.Web.Services;

/// <summary>
/// Model Info › Documents: what an uploaded workbook is, so several files picked at once each go to their place —
/// a JIG / Machine / Tester list (its JIG LIST, Machine List or Tester List sheets), Work Instructions (a sheet per
/// process), or neither (with why, when the workbook could not be opened, e.g. still encrypted).
/// </summary>
public static class WorkingDocumentKind
{
    public const string JigList = "jig-list", Instructions = "instructions", Unknown = "unknown";

    /// <summary>The largest workbook it reads: the larger of the two kinds (Work Instructions with their pictures).</summary>
    public const long MaxUploadBytes = WorkingInstructionService.MaxUploadBytes > WorkingProcessListService.MaxUploadBytes
        ? WorkingInstructionService.MaxUploadBytes : WorkingProcessListService.MaxUploadBytes;

    public sealed record Found(string Kind, string Reason);

    public static Found Of(Stream xlsx)
    {
        string tmp = Path.Combine(Path.GetTempPath(), $"doc-kind-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (FileStream f = File.Create(tmp)) xlsx.CopyTo(f);
            try
            {
                using FileStream f = File.OpenRead(tmp);
                if (WorkingProcessListService.ReadItems(f).Count > 0) return new(JigList, "");
            }
            catch (ArgumentException) { }
            try { return WorkingInstructionService.ParseFile(tmp).Processes.Count > 0 ? new(Instructions, "") : new(Unknown, ""); }
            catch (ArgumentException ex) { return new(Unknown, ex.Message); }
        }
        finally { try { File.Delete(tmp); } catch { } }
    }
}
