using System.Collections.Generic;
using System.Windows;

namespace GraphMaker
{
    public partial class ValuePlotFileSettingsWindow : Window
    {
        public string Delimiter { get; private set; } = "\t";
        public int HeaderRowNumber { get; private set; } = 1;
        public bool IsXAxisDate { get; private set; } = true;
        public int DataColorIndex { get; private set; }
        public int SpecColorIndex { get; private set; }
        public int UpperColorIndex { get; private set; }
        public int LowerColorIndex { get; private set; }
        public string SpecValue { get; private set; } = string.Empty;
        public string UpperValue { get; private set; } = string.Empty;
        public string LowerValue { get; private set; } = string.Empty;

        public ValuePlotFileSettingsWindow(FileInfo_DailySampling fileInfo, IReadOnlyList<string> colorNames)
        {
            InitializeComponent();

            DataColorComboBox.ItemsSource = colorNames;
            SpecColorComboBox.ItemsSource = colorNames;
            UpperColorComboBox.ItemsSource = colorNames;
            LowerColorComboBox.ItemsSource = colorNames;

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
            XAxisDateRadio.IsChecked = fileInfo.IsXAxisDate;
            XAxisSequenceRadio.IsChecked = !fileInfo.IsXAxisDate;

            DataColorComboBox.SelectedIndex = fileInfo.SavedDataColorIndex;
            SpecColorComboBox.SelectedIndex = fileInfo.SavedSpecColorIndex;
            UpperColorComboBox.SelectedIndex = fileInfo.SavedUpperColorIndex;
            LowerColorComboBox.SelectedIndex = fileInfo.SavedLowerColorIndex;

            SpecValueTextBox.Text = fileInfo.SavedSpecValue;
            UpperLimitValueTextBox.Text = fileInfo.SavedUpperValue;
            LowerLimitValueTextBox.Text = fileInfo.SavedLowerValue;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out var headerRow) || headerRow <= 0)
            {
                MessageBox.Show("Header row must be a positive integer.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Delimiter = TabDelimiterRadio.IsChecked == true ? "\t" :
                CommaDelimiterRadio.IsChecked == true ? "," : " ";
            HeaderRowNumber = headerRow;
            IsXAxisDate = XAxisDateRadio.IsChecked == true;
            DataColorIndex = DataColorComboBox.SelectedIndex < 0 ? 0 : DataColorComboBox.SelectedIndex;
            SpecColorIndex = SpecColorComboBox.SelectedIndex < 0 ? 0 : SpecColorComboBox.SelectedIndex;
            UpperColorIndex = UpperColorComboBox.SelectedIndex < 0 ? 0 : UpperColorComboBox.SelectedIndex;
            LowerColorIndex = LowerColorComboBox.SelectedIndex < 0 ? 0 : LowerColorComboBox.SelectedIndex;
            SpecValue = SpecValueTextBox.Text?.Trim() ?? string.Empty;
            UpperValue = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty;
            LowerValue = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty;

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
