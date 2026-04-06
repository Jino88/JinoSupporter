using System.Windows;
using System.Windows.Controls;

namespace GraphMaker
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ScatterPlotButton_Click(object sender, RoutedEventArgs e)
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            ContentArea.Content = new ScatterPlotView();
        }

        private void ValuePlotButton_Click(object sender, RoutedEventArgs e)
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            ContentArea.Content = new ValuePlotView();
        }

        private void UnifiedMultiYButton_Click(object sender, RoutedEventArgs e)
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            ContentArea.Content = new UnifiedMultiYView();
        }

        private void HeatMapButton_Click(object sender, RoutedEventArgs e)
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            ContentArea.Content = new HeatMapView();
        }

        private void AudioBusDataButton_Click(object sender, RoutedEventArgs e)
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            ContentArea.Content = new AudioBusDataView();
        }
    }
}
