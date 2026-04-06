using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace GraphMaker
{
    public partial class ProcessTrendPairChoiceWindow : Window
    {
        public ProcessTrendComputationCandidate? SelectedCandidate { get; private set; }

        public ProcessTrendPairChoiceWindow(string groupTitle, IReadOnlyList<ProcessTrendComputationCandidate> candidates)
        {
            InitializeComponent();

            HeaderTextBlock.Text = $"Select pair for calculation ({groupTitle})";
            PairListBox.ItemsSource = candidates.ToList();
            PairListBox.SelectedIndex = candidates.Count > 0 ? 0 : -1;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedCandidate = PairListBox.SelectedItem as ProcessTrendComputationCandidate;
            if (SelectedCandidate == null)
            {
                MessageBox.Show("Select a pair first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
