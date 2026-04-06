namespace GraphMaker
{
    public enum XAxisMode
    {
        Date = 0,
        Sequence = 1,
        None = 2
    }

    public enum NoXAxisDisplayMode
    {
        SampleOrderOnXAxis = 0,
        FixedSingleX = 1
    }

    public enum ValuePlotDisplayMode
    {
        Combined = 0,
        SplitByColumn = 1
    }

    public class ColumnLimitSetting
    {
        public string ColumnName { get; set; } = string.Empty;
        public string SpecValue { get; set; } = string.Empty;
        public string UpperValue { get; set; } = string.Empty;
        public string LowerValue { get; set; } = string.Empty;
    }
}
