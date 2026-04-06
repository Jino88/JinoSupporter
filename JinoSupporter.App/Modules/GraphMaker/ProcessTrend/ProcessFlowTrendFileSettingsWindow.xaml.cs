using System.Collections.Generic;
using System.Windows;

namespace GraphMaker
{
    public partial class ProcessFlowTrendFileSettingsWindow : Window
    {
        public string Delimiter { get; private set; } = "\t";
        public int HeaderRowNumber { get; private set; } = 1;
        public bool UseFirstColumnAsSampleId { get; private set; }
        public int MaxSamples { get; private set; } = 100;
        public int PlotColorIndex { get; private set; }
        public bool UseQuadraticRegression { get; private set; }

        public ProcessFlowTrendFileSettingsWindow(ProcessTrendFileInfo fileInfo, IReadOnlyList<string> colorNames)
        {
            InitializeComponent();

            PlotColorComboBox.ItemsSource = colorNames;

            if (fileInfo.Delimiter == ",")
            {
                CommaDelimiterRadio.IsChecked = true;
            }
            else if (fileInfo.Delimiter == " ")
            {
                SpaceDelimiterRadio.IsChecked = true;
            }
            else
            {
                TabDelimiterRadio.IsChecked = true;
            }

            HeaderRowTextBox.Text = fileInfo.HeaderRowNumber.ToString();
            UseFirstColumnCheckBox.IsChecked = fileInfo.UseFirstColumnAsSampleId;
            MaxSamplesTextBox.Text = fileInfo.MaxSamples.ToString();
            PlotColorComboBox.SelectedIndex = fileInfo.SavedPlotColorIndex;
            LinearDegreeRadio.IsChecked = !fileInfo.UseQuadraticRegression;
            QuadraticDegreeRadio.IsChecked = fileInfo.UseQuadraticRegression;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out var headerRow) || headerRow <= 0)
            {
                MessageBox.Show("Header row must be a positive integer.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(MaxSamplesTextBox.Text, out var maxSamples) || maxSamples <= 0)
            {
                MessageBox.Show("Max samples must be a positive integer.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Delimiter = TabDelimiterRadio.IsChecked == true ? "\t" :
                CommaDelimiterRadio.IsChecked == true ? "," : " ";
            HeaderRowNumber = headerRow;
            UseFirstColumnAsSampleId = UseFirstColumnCheckBox.IsChecked == true;
            MaxSamples = maxSamples;
            PlotColorIndex = PlotColorComboBox.SelectedIndex < 0 ? 0 : PlotColorComboBox.SelectedIndex;
            UseQuadraticRegression = QuadraticDegreeRadio.IsChecked == true;
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
