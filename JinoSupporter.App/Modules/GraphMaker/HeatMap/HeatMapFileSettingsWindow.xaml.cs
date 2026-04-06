using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace GraphMaker
{
    public partial class HeatMapFileSettingsWindow : Window
    {
        private readonly string _filePath;
        private List<string> _headers = new();

        public string Delimiter { get; private set; } = "\t";
        public int HeaderRowNumber { get; private set; } = 1;
        public int Condition1Index { get; private set; }
        public int Condition2Index { get; private set; }
        public int ExtraCondition1Index { get; private set; } = -1;
        public int ExtraCondition2Index { get; private set; } = -1;
        public int ResultIndex { get; private set; }
        public int NgColumnIndex { get; private set; }
        public int InputColumnIndex { get; private set; }
        public bool UseFormulaResult { get; private set; }

        public HeatMapFileSettingsWindow(string filePath, HeatMapFileInfo fileInfo)
        {
            InitializeComponent();
            _filePath = filePath;

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
            UseFormulaResult = fileInfo.UseFormulaResult;
            LoadHeadersAndBind(
                fileInfo.SavedCondition1Index,
                fileInfo.SavedCondition2Index,
                fileInfo.SavedExtraCondition1Index,
                fileInfo.SavedExtraCondition2Index,
                fileInfo.SavedResultIndex,
                fileInfo.SavedNgColumnIndex,
                fileInfo.SavedInputColumnIndex);
        }

        private void ReloadColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            LoadHeadersAndBind(
                Condition1ComboBox.SelectedIndex,
                Condition2ComboBox.SelectedIndex,
                ExtraCondition1ComboBox.SelectedIndex - 1,
                ExtraCondition2ComboBox.SelectedIndex - 1,
                ResultColumnComboBox.SelectedIndex,
                NgColumnComboBox.SelectedIndex,
                InputColumnComboBox.SelectedIndex);
        }

        private void LoadHeadersAndBind(
            int condition1Index,
            int condition2Index,
            int extraCondition1Index,
            int extraCondition2Index,
            int resultIndex,
            int ngColumnIndex,
            int inputColumnIndex)
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out var headerRow) || headerRow <= 0)
            {
                MessageBox.Show("Header row must be a positive integer.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var delimiter = GetDelimiterFromUi();
            var headers = ReadHeaders(_filePath, delimiter, headerRow);
            if (headers.Count == 0)
            {
                MessageBox.Show("Could not read header row with current settings.", "Invalid Data",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _headers = headers;
            BindHeaderCombos();

            SetComboIndex(Condition1ComboBox, condition1Index, 0);
            SetComboIndex(Condition2ComboBox, condition2Index, Math.Min(1, _headers.Count - 1));
            SetComboIndex(ExtraCondition1ComboBox, extraCondition1Index + 1, 0);
            SetComboIndex(ExtraCondition2ComboBox, extraCondition2Index + 1, 0);
            SetComboIndex(ResultColumnComboBox, resultIndex, _headers.Count > 2 ? 2 : _headers.Count - 1);
            SetComboIndex(NgColumnComboBox, ngColumnIndex, 0);
            SetComboIndex(InputColumnComboBox, inputColumnIndex, Math.Min(1, _headers.Count - 1));

            if (UseFormulaResult)
            {
                FormulaResultRadio.IsChecked = true;
            }
            else
            {
                DirectResultRadio.IsChecked = true;
            }

            ApplyResultModeState();
        }

        private void BindHeaderCombos()
        {
            Condition1ComboBox.ItemsSource = _headers;
            Condition2ComboBox.ItemsSource = _headers;
            ResultColumnComboBox.ItemsSource = _headers;
            NgColumnComboBox.ItemsSource = _headers;
            InputColumnComboBox.ItemsSource = _headers;

            var extraHeaders = new List<string> { "(none)" };
            extraHeaders.AddRange(_headers);
            ExtraCondition1ComboBox.ItemsSource = extraHeaders;
            ExtraCondition2ComboBox.ItemsSource = extraHeaders;
        }

        private static void SetComboIndex(System.Windows.Controls.ComboBox comboBox, int index, int fallback)
        {
            if (comboBox.Items.Count == 0)
            {
                return;
            }

            if (index >= 0 && index < comboBox.Items.Count)
            {
                comboBox.SelectedIndex = index;
                return;
            }

            comboBox.SelectedIndex = fallback >= 0 && fallback < comboBox.Items.Count ? fallback : 0;
        }

        private static List<string> ReadHeaders(string filePath, string delimiter, int headerRowNumber)
        {
            var lines = File.ReadAllLines(filePath);
            var headerIndex = headerRowNumber - 1;
            if (headerIndex < 0 || headerIndex >= lines.Length)
            {
                return new List<string>();
            }

            var rawHeaders = GraphMakerTableHelper.SplitLine(lines[headerIndex], delimiter);
            return GraphMakerTableHelper.BuildUniqueHeaders(rawHeaders);
        }

        private void ResultModeChanged(object sender, RoutedEventArgs e)
        {
            UseFormulaResult = FormulaResultRadio.IsChecked == true;
            ApplyResultModeState();
        }

        private void ApplyResultModeState()
        {
            ResultColumnComboBox.IsEnabled = !UseFormulaResult;
            NgColumnComboBox.IsEnabled = UseFormulaResult;
            InputColumnComboBox.IsEnabled = UseFormulaResult;
        }

        private string GetDelimiterFromUi()
        {
            return TabDelimiterRadio.IsChecked == true ? "\t" :
                CommaDelimiterRadio.IsChecked == true ? "," : " ";
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (_headers.Count == 0)
            {
                MessageBox.Show("No columns loaded. Check delimiter/header row.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(HeaderRowTextBox.Text, out var headerRow) || headerRow <= 0)
            {
                MessageBox.Show("Header row must be a positive integer.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (Condition1ComboBox.SelectedIndex < 0 || Condition2ComboBox.SelectedIndex < 0)
            {
                MessageBox.Show("Condition 1 and Condition 2 are required.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!UseFormulaResult && ResultColumnComboBox.SelectedIndex < 0)
            {
                MessageBox.Show("Result column is required in direct mode.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (UseFormulaResult && (NgColumnComboBox.SelectedIndex < 0 || InputColumnComboBox.SelectedIndex < 0))
            {
                MessageBox.Show("NG and Input columns are required in formula mode.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Delimiter = GetDelimiterFromUi();
            HeaderRowNumber = headerRow;
            Condition1Index = Condition1ComboBox.SelectedIndex;
            Condition2Index = Condition2ComboBox.SelectedIndex;
            ExtraCondition1Index = ExtraCondition1ComboBox.SelectedIndex - 1;
            ExtraCondition2Index = ExtraCondition2ComboBox.SelectedIndex - 1;
            ResultIndex = ResultColumnComboBox.SelectedIndex;
            NgColumnIndex = NgColumnComboBox.SelectedIndex;
            InputColumnIndex = InputColumnComboBox.SelectedIndex;
            UseFormulaResult = FormulaResultRadio.IsChecked == true;

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
