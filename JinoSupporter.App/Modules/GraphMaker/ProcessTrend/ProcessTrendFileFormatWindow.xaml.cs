using System.Windows;

namespace GraphMaker
{
    public partial class ProcessTrendFileFormatWindow : Window
    {
        public string SelectedDelimiter { get; private set; } = "\t";
        public int HeaderRowNumber { get; private set; } = 1;

        public ProcessTrendFileFormatWindow()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out int headerRow) || headerRow <= 0)
            {
                MessageBox.Show("Header Row must be a positive number.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            HeaderRowNumber = headerRow;
            if (CommaDelimiterRadio.IsChecked == true)
            {
                SelectedDelimiter = ",";
            }
            else if (SpaceDelimiterRadio.IsChecked == true)
            {
                SelectedDelimiter = " ";
            }
            else
            {
                SelectedDelimiter = "\t";
            }

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
