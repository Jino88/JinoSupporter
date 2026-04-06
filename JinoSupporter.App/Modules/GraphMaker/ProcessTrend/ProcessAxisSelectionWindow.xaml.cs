using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace GraphMaker
{
    public partial class ProcessAxisSelectionWindow : Window
    {
        private sealed class XAxisEntry
        {
            public int XAxisProcessIndex { get; init; }
            public string Name { get; init; } = string.Empty;
            public Dictionary<int, int> YAxisGroups { get; set; } = new();
            public override string ToString()
            {
                if (YAxisGroups.Count == 0)
                {
                    return $"X: {Name}";
                }

                var yNames = YAxisGroups.Keys
                    .OrderBy(index => index)
                    .Select(index => index >= 0 && index < _owner._processNames.Count ? _owner._processNames[index] : "?")
                    .ToList();
                return $"X: {Name} -> Y: {string.Join(", ", yNames)}";
            }

            public required ProcessAxisSelectionWindow _owner { get; init; }
        }

        private sealed class YAxisOption
        {
            public int Index { get; init; }
            public CheckBox CheckBox { get; init; } = null!;
        }

        private sealed class SelectedYGroupRow
        {
            public int YIndex { get; init; }
            public string Label { get; init; } = string.Empty;
            public ComboBox ComboBox { get; init; } = null!;
        }

        private readonly List<string> _processNames;
        private readonly List<YAxisOption> _yAxisOptions = new();
        private readonly List<XAxisEntry> _entries = new();
        private readonly Dictionary<int, string> _groupOptions = new();
        private readonly List<SelectedYGroupRow> _selectedYGroupRows = new();
        private bool _isUpdatingUi;

        public IReadOnlyList<ProcessAxisMapping> AxisMappings =>
            _entries
                .SelectMany(entry => entry.YAxisGroups
                    .GroupBy(pair => pair.Value)
                    .Select(group => new ProcessAxisMapping
                    {
                        XAxisProcessIndex = entry.XAxisProcessIndex,
                        YAxisProcessIndices = group.Select(pair => pair.Key).OrderBy(index => index).ToList(),
                        GroupId = group.Key
                    }))
                .Where(mapping => mapping.YAxisProcessIndices.Count > 0)
                .ToList();
        public IReadOnlyDictionary<int, string> GroupOptions => _groupOptions;

        public ProcessAxisSelectionWindow(
            IReadOnlyList<string> processNames,
            IReadOnlyList<ProcessAxisMapping> savedMappings,
            IReadOnlyDictionary<int, string> groupOptions)
        {
            InitializeComponent();
            _processNames = processNames.ToList();

            if (groupOptions.Count > 0)
            {
                foreach (var group in groupOptions.OrderBy(g => g.Key))
                {
                    if (group.Key > 0)
                    {
                        _groupOptions[group.Key] = string.IsNullOrWhiteSpace(group.Value) ? $"Group {group.Key}" : group.Value;
                    }
                }
            }
            if (_groupOptions.Count == 0)
            {
                _groupOptions[1] = "Group 1";
                _groupOptions[2] = "Group 2";
            }

            var savedByX = savedMappings
                .GroupBy(mapping => mapping.XAxisProcessIndex)
                .ToDictionary(
                    group => group.Key,
                    group => group.SelectMany(mapping => mapping.YAxisProcessIndices.Select(y => new { Y = y, mapping.GroupId }))
                        .GroupBy(item => item.Y)
                        .ToDictionary(item => item.Key, item => item.Last().GroupId));
            for (int i = 0; i < processNames.Count; i++)
            {
                var saved = savedByX.TryGetValue(i, out var existing) ? existing : null;
                _entries.Add(new XAxisEntry
                {
                    _owner = this,
                    XAxisProcessIndex = i,
                    Name = processNames[i],
                    YAxisGroups = saved ?? new Dictionary<int, int>()
                });

                var checkBox = new CheckBox
                {
                    Content = processNames[i],
                    Margin = new Thickness(0, 0, 0, 6)
                };
                checkBox.Checked += YAxisCheckBox_Changed;
                checkBox.Unchecked += YAxisCheckBox_Changed;
                YAxisItemsPanel.Children.Add(checkBox);
                _yAxisOptions.Add(new YAxisOption { Index = i, CheckBox = checkBox });
            }

            if (savedMappings.Count == 0 && processNames.Count > 1)
            {
                foreach (int index in Enumerable.Range(1, processNames.Count - 1))
                {
                    _entries[0].YAxisGroups[index] = _groupOptions.Keys.Min();
                }
            }

            RefreshGroupOptions(_groupOptions.Keys.Min());
            XAxisListBox.ItemsSource = _entries;
            if (XAxisListBox.Items.Count > 0)
            {
                XAxisListBox.SelectedIndex = 0;
            }
        }

        private void XAxisListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (XAxisListBox.SelectedItem is XAxisEntry selected)
            {
                LoadEntryToUi(selected);
            }
        }

        private void YAxisCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingUi || XAxisListBox.SelectedItem is not XAxisEntry selected)
            {
                return;
            }

            var checkedIndices = _yAxisOptions
                .Where(option => option.CheckBox.IsChecked == true && option.Index != selected.XAxisProcessIndex)
                .Select(option => option.Index)
                .Distinct()
                .OrderBy(index => index)
                .ToList();

            var preserved = new Dictionary<int, int>();
            foreach (int index in checkedIndices)
            {
                preserved[index] = selected.YAxisGroups.TryGetValue(index, out int groupId)
                    ? groupId
                    : _groupOptions.Keys.Min();
            }

            selected.YAxisGroups = preserved;
            RefreshSelectedYGroupList();
            XAxisListBox.Items.Refresh();
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var option in _yAxisOptions)
            {
                option.CheckBox.IsChecked = true;
            }
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var option in _yAxisOptions)
            {
                option.CheckBox.IsChecked = false;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var mapping in _entries)
            {
                mapping.YAxisGroups = mapping.YAxisGroups
                    .Where(pair => pair.Key >= 0 && pair.Key < _processNames.Count && pair.Key != mapping.XAxisProcessIndex)
                    .GroupBy(pair => pair.Key)
                    .ToDictionary(group => group.Key, group => group.Last().Value);
            }

            if (!_entries.Any(entry => entry.YAxisGroups.Count > 0))
            {
                MessageBox.Show("Select at least one Y-axis column for at least one X-axis.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void AddGroupButton_Click(object sender, RoutedEventArgs e)
        {
            int nextId = _groupOptions.Count == 0 ? 1 : _groupOptions.Keys.Max() + 1;
            _groupOptions[nextId] = $"Group {nextId}";
            RefreshGroupOptions(nextId);
            RefreshSelectedYGroupList();
        }

        private void RemoveGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (_groupOptions.Count <= 1 || GroupManageComboBox.SelectedValue is not int groupId || !_groupOptions.ContainsKey(groupId))
            {
                return;
            }

            int fallback = _groupOptions.Keys.First(key => key != groupId);
            _groupOptions.Remove(groupId);
            foreach (var entry in _entries)
            {
                foreach (int yIndex in entry.YAxisGroups.Where(pair => pair.Value == groupId).Select(pair => pair.Key).ToList())
                {
                    entry.YAxisGroups[yIndex] = fallback;
                }
            }

            RefreshGroupOptions(fallback);
            RefreshSelectedYGroupList();
            XAxisListBox.Items.Refresh();
        }

        private void ApplyGroupNameButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupManageComboBox.SelectedValue is not int groupId || !_groupOptions.ContainsKey(groupId))
            {
                return;
            }

            string name = (GroupNameTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("Group name cannot be empty.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _groupOptions[groupId] = name;
            RefreshGroupOptions(groupId);
            RefreshSelectedYGroupList();
            XAxisListBox.Items.Refresh();
        }

        private void LoadEntryToUi(XAxisEntry? entry)
        {
            _isUpdatingUi = true;
            try
            {
                RefreshGroupOptions(_groupOptions.Keys.Min());
                foreach (var option in _yAxisOptions)
                {
                    option.CheckBox.IsChecked = entry != null &&
                                                option.Index != entry.XAxisProcessIndex &&
                                                entry.YAxisGroups.ContainsKey(option.Index);
                    option.CheckBox.IsEnabled = entry != null && option.Index != entry.XAxisProcessIndex;
                }
            }
            finally
            {
                _isUpdatingUi = false;
            }

            RefreshSelectedYGroupList();
        }

        private void RefreshGroupOptions(int selectedGroupId)
        {
            _isUpdatingUi = true;
            try
            {
                GroupManageComboBox.ItemsSource = null;
                GroupManageComboBox.ItemsSource = _groupOptions.OrderBy(g => g.Key).ToList();
                GroupManageComboBox.SelectedValue = selectedGroupId;
                GroupNameTextBox.Text = _groupOptions.TryGetValue(selectedGroupId, out string? name) ? name : string.Empty;
            }
            finally
            {
                _isUpdatingUi = false;
            }
        }

        private void RefreshSelectedYGroupList()
        {
            _selectedYGroupRows.Clear();
            SelectedYGroupListBox.Items.Clear();

            if (_entries.Count == 0)
            {
                return;
            }

            foreach (var entry in _entries.OrderBy(item => item.XAxisProcessIndex))
            {
                foreach (var pair in entry.YAxisGroups.OrderBy(pair => pair.Key))
                {
                    var rowGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });

                    var textBlock = new TextBlock
                    {
                        Text = $"X: {entry.Name} -> Y: {_processNames[pair.Key]}",
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    var comboBox = new ComboBox
                    {
                        ItemsSource = _groupOptions.OrderBy(g => g.Key).ToList(),
                        DisplayMemberPath = "Value",
                        SelectedValuePath = "Key",
                        SelectedValue = pair.Value
                    };
                    int yIndex = pair.Key;
                    comboBox.SelectionChanged += (_, _) =>
                    {
                        if (_isUpdatingUi || comboBox.SelectedValue is not int selectedGroupId)
                        {
                            return;
                        }

                        entry.YAxisGroups[yIndex] = selectedGroupId;
                        XAxisListBox.Items.Refresh();
                    };

                    Grid.SetColumn(textBlock, 0);
                    Grid.SetColumn(comboBox, 1);
                    rowGrid.Children.Add(textBlock);
                    rowGrid.Children.Add(comboBox);
                    SelectedYGroupListBox.Items.Add(rowGrid);
                }
            }
        }
    }
}
