using System.Collections.Generic;
using System.Linq;
using System.Windows;
using OxyPlot;

namespace GraphMaker
{
    public sealed class ProcessTrendComputationCandidate
    {
        public string PairTitle { get; init; } = string.Empty;
        public PlotModel PlotModel { get; init; } = new();
        public string XAxisTitle { get; init; } = "X";
        public string YAxisTitle { get; init; } = "Y";
        public IReadOnlyList<DataPoint> RawPoints { get; init; } = new List<DataPoint>();
    }

    public sealed class ProcessPairPlotResult
    {
        public string PairTitle { get; init; } = string.Empty;
        public PlotModel PlotModel { get; init; } = new();
        public string DetailText { get; init; } = string.Empty;
        public string XAxisTitle { get; init; } = "X";
        public string YAxisTitle { get; init; } = "Y";
        public IReadOnlyList<DataPoint> RawPoints { get; init; } = new List<DataPoint>();
        public IReadOnlyList<ProcessTrendComputationCandidate> ComputationCandidates { get; init; } =
            new List<ProcessTrendComputationCandidate>();
    }

    public partial class ProcessFlowTrendResultWindow : Window
    {
        public List<ProcessPairPlotResult> PairResults { get; }

        public ProcessFlowTrendResultWindow(IReadOnlyList<ProcessPairPlotResult> pairResults, string windowTitle = "Process Trend Result")
        {
            InitializeComponent();
            Title = windowTitle;
            PairResults = pairResults.ToList();
            DataContext = this;
        }

        private void PairTitle_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is not FrameworkElement element || element.DataContext is not ProcessPairPlotResult pair)
            {
                return;
            }

            OpenSingleGraphWindow(pair);
        }

        private void OpenSingleGraphWindow(ProcessPairPlotResult pair)
        {
            var window = new ProcessTrendLargeDetailWindow(pair)
            {
                WindowState = WindowState.Maximized,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this
            };
            window.Show();
        }
    }
}
