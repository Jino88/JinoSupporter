namespace JinoSupporter.Web.Services;

public sealed class Test3EditorSnapshot
{
    public List<string> ModelNames { get; init; } = [];
    public string SelectedModel { get; init; } = string.Empty;
    public List<Test3EditorSide> Sides { get; init; } = [];
    public List<Test3EditorMaterial> Materials { get; init; } = [];
    public bool BomBusy { get; init; }
    public string BomSource { get; init; } = string.Empty;
    public string BomError { get; init; } = string.Empty;
    public string StatusMessage { get; init; } = string.Empty;
    public bool StatusError { get; init; }
}

public sealed class Test3EditorSide
{
    public string ModelName { get; init; } = string.Empty;
    public string SideLabel { get; init; } = string.Empty;

    /// <summary>L/R 짝 모델. 값이 있으면 이 보드의 편집이 짝 모델에도 그대로 반영된다.</summary>
    public string MirrorModelName { get; init; } = string.Empty;

    public string MirrorSideLabel { get; init; } = string.Empty;
    public List<Test3EditorLane> Lanes { get; init; } = [];
    public List<Test3EditorProcess> AvailableProcesses { get; init; } = [];
}

public sealed class Test3EditorLane
{
    public string LaneCode { get; init; } = "MAIN";
    public string MergeTargetProcessId { get; init; } = string.Empty;
    public string MergeTargetLabel { get; init; } = string.Empty;
    public List<Test3EditorProcess> Processes { get; init; } = [];
}

public sealed class Test3EditorProcess
{
    public string Id { get; init; } = string.Empty;
    public string ModelName { get; init; } = string.Empty;
    public string ProcessCode { get; init; } = string.Empty;
    public string ProcessName { get; init; } = string.Empty;
    public string ProcessType { get; init; } = string.Empty;
    public string LaneCode { get; init; } = string.Empty;
    public string ProcessNo { get; init; } = string.Empty;
    public int Order { get; init; }

    /// <summary>짝 모델에서 같이 배치되는 공정 코드. 짝이 없으면 빈 문자열.</summary>
    public string MirrorProcessCode { get; init; } = string.Empty;
}

public sealed class Test3EditorMaterial
{
    public string Id { get; init; } = string.Empty;
    public string ModelName { get; init; } = string.Empty;
    public string SideLabel { get; init; } = string.Empty;
    public string MaterialCode { get; init; } = string.Empty;
    public string MaterialName { get; init; } = string.Empty;
    public decimal UsageQty { get; init; }
    public string UsageUnit { get; init; } = "PC";
    public string AssignedProcessId { get; init; } = string.Empty;
    public string AssignedProcessLabel { get; init; } = string.Empty;
    public string ScopeLabel { get; init; } = string.Empty;

    /// <summary>짝 모델에서 대신 투입되는 자재(-L → -R 치환 결과). 짝이 없으면 빈 문자열.</summary>
    public string MirrorMaterialCode { get; init; } = string.Empty;

    public string MirrorMaterialName { get; init; } = string.Empty;
}

public sealed class Test3EditorLayoutRequest
{
    public List<Test3EditorLaneRequest> Lanes { get; init; } = [];
}

public sealed class Test3EditorLaneRequest
{
    public string ModelName { get; init; } = string.Empty;
    public string LaneCode { get; init; } = "MAIN";
    public string MergeTargetProcessId { get; init; } = string.Empty;
    public List<string> ProcessIds { get; init; } = [];
}

public sealed class Test3EditorMaterialRequest
{
    public string MaterialId { get; init; } = string.Empty;
    public string ProcessId { get; init; } = string.Empty;
    public decimal UsageQty { get; init; }
    public string UsageUnit { get; init; } = "PC";
}
