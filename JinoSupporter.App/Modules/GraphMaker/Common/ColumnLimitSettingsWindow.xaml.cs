using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace GraphMaker
{
    public partial class ColumnLimitSettingsWindow : Window
    {
        private readonly ObservableCollection<ColumnLimitSetting> _items = new();

        public Dictionary<string, ColumnLimitSetting> ResultLimits { get; private set; } =
            new Dictionary<string, ColumnLimitSetting>();

        public ColumnLimitSettingsWindow(
            IEnumerable<string> columnNames,
            IDictionary<string, ColumnLimitSetting>? existingLimits)
        {
            InitializeComponent();

            foreach (var name in columnNames)
            {
                if (existingLimits != null && existingLimits.TryGetValue(name, out var existing))
                {
                    _items.Add(new ColumnLimitSetting
                    {
                        ColumnName = name,
                        SpecValue = existing.SpecValue,
                        UpperValue = existing.UpperValue,
                        LowerValue = existing.LowerValue
                    });
                }
                else
                {
                    _items.Add(new ColumnLimitSetting { ColumnName = name });
                }
            }

            ColumnLimitDataGrid.ItemsSource = _items;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ResultLimits = _items.ToDictionary(
                x => x.ColumnName,
                x => new ColumnLimitSetting
                {
                    ColumnName = x.ColumnName,
                    SpecValue = x.SpecValue?.Trim() ?? string.Empty,
                    UpperValue = x.UpperValue?.Trim() ?? string.Empty,
                    LowerValue = x.LowerValue?.Trim() ?? string.Empty
                });

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
