using DiskTree.Services;
using System.Collections.ObjectModel;

namespace DiskTree.Models;

public sealed class DiskNode
{
    public string Name { get; init; } = string.Empty;
    public string FullPath { get; init; } = string.Empty;
    public bool IsDirectory { get; init; }
    public long SizeBytes { get; set; }
    public double PercentOfRoot { get; set; }
    public ObservableCollection<DiskNode> Children { get; } = new();

    public string SizeText => ByteSizeFormatter.Format(SizeBytes);
    public string PercentText => $"{PercentOfRoot:F2}%";
    public string KindText => IsDirectory ? "DIR" : "FILE";
}
