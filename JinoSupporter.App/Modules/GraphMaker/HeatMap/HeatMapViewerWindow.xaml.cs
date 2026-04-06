using System.Windows;
using OxyPlot;

namespace GraphMaker
{
    public partial class HeatMapViewerWindow : Window
    {
        public PlotModel HeatMapModel { get; }

        public HeatMapViewerWindow(PlotModel heatMapModel)
        {
            InitializeComponent();
            HeatMapModel = heatMapModel;
            DataContext = this;

            if (!string.IsNullOrWhiteSpace(heatMapModel.Title))
            {
                Title = heatMapModel.Title;
            }
        }
    }
}
