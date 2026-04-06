using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Border = System.Windows.Controls.Border;
using ComboBox = System.Windows.Controls.ComboBox;
using ListBox = System.Windows.Controls.ListBox;
using ListBoxItem = System.Windows.Controls.ListBoxItem;
using TextBox = System.Windows.Controls.TextBox;

namespace GraphMaker
{
    public partial class ScatterFileSettingsWindow : Window
    {
        private sealed class ColumnItem
        {
            public int Index { get; set; }
            public string Header { get; set; } = string.Empty;
            public override string ToString() => Header;
        }

        private sealed class GroupRow
        {
            public int GroupId { get; set; }
            public Border Container { get; set; } = null!;
            public TextBox NameTextBox { get; set; } = null!;
            public ComboBox ColorComboBox { get; set; } = null!;
            public ComboBox LineStyleComboBox { get; set; } = null!;
            public ListBox GroupColumnsListBox { get; set; } = null!;
        }

        private readonly FileInfo_Custom _fileInfo;
        private readonly Dictionary<int, int> _columnGroups = new Dictionary<int, int>();
        private readonly List<GroupRow> _groupRows = new List<GroupRow>();
        private List<string> _headers = new List<string>();
        private int _nextGroupId = 1;

        private readonly List<ColorOption> _colorOptions = new List<ColorOption>
        {
            new ColorOption { Name = "Black", Color = Colors.Black },
            new ColorOption { Name = "Red", Color = Colors.Red },
            new ColorOption { Name = "Blue", Color = Colors.Blue },
            new ColorOption { Name = "Green", Color = Colors.Green },
            new ColorOption { Name = "Orange", Color = Colors.Orange },
            new ColorOption { Name = "Purple", Color = Colors.Purple },
            new ColorOption { Name = "Gray", Color = Colors.Gray }
        };

        private readonly List<string> _lineStyleOptions = new List<string> { "Solid", "Dash", "Dot", "DashDot" };

        public ScatterFileSettingsWindow(FileInfo_Custom fileInfo)
        {
            InitializeComponent();
            _fileInfo = fileInfo;
            foreach (var kv in fileInfo.SavedColumnGroups)
            {
                _columnGroups[kv.Key] = kv.Value;
            }

            EnsureSavedGroups();
            InitializeFromFileSettings();
            ReloadColumns();
        }

        private void EnsureSavedGroups()
        {
            if (_fileInfo.SavedGroups == null)
            {
                _fileInfo.SavedGroups = new List<SavedGroupSetting>();
            }

            if (_fileInfo.SavedGroups.Count == 0)
            {
                _fileInfo.SavedGroups.Add(new SavedGroupSetting { GroupId = 1, Name = "Group 1", ColorIndex = 2, LineStyleIndex = 0 });
                _fileInfo.SavedGroups.Add(new SavedGroupSetting { GroupId = 2, Name = "Group 2", ColorIndex = 1, LineStyleIndex = 0 });
            }
        }

        private void InitializeFromFileSettings()
        {
            if (_fileInfo.Delimiter == ",")
            {
                CommaDelimiterRadio.IsChecked = true;
            }
            else if (_fileInfo.Delimiter == " ")
            {
                SpaceDelimiterRadio.IsChecked = true;
            }
            else
            {
                TabDelimiterRadio.IsChecked = true;
            }

            HeaderRowTextBox.Text = _fileInfo.HeaderRowNumber > 0 ? _fileInfo.HeaderRowNumber.ToString() : "1";
            XAxisModeComboBox.SelectedIndex = _fileInfo.SavedXAxisIndex > 0 ? 0 : 1;
            XAxisLogScaleCheckBox.IsChecked = _fileInfo.SavedXAxisLogScale;

            BindColorCombo(SpecColorComboBox, _fileInfo.SavedSpecColorIndex);
            BindColorCombo(UpperLimitColorComboBox, _fileInfo.SavedUpperColorIndex);
            BindColorCombo(LowerLimitColorComboBox, _fileInfo.SavedLowerColorIndex);
            BindLineStyleCombo(SpecLineStyleComboBox, _fileInfo.SavedSpecLineStyleIndex);
            BindLineStyleCombo(UpperLimitLineStyleComboBox, _fileInfo.SavedUpperLineStyleIndex);
            BindLineStyleCombo(LowerLimitLineStyleComboBox, _fileInfo.SavedLowerLineStyleIndex);

            BuildGroupRowsFromSavedGroups();
        }

        private void BuildGroupRowsFromSavedGroups()
        {
            GroupRowsHost.Children.Clear();
            _groupRows.Clear();

            _nextGroupId = _fileInfo.SavedGroups.Max(g => g.GroupId) + 1;
            foreach (var savedGroup in _fileInfo.SavedGroups.OrderBy(g => g.GroupId))
            {
                AddGroupRow(savedGroup.GroupId, savedGroup.Name, savedGroup.ColorIndex, savedGroup.LineStyleIndex);
            }
        }

        private void AddGroupButton_Click(object sender, RoutedEventArgs e)
        {
            int groupId = _nextGroupId++;
            AddGroupRow(groupId, $"Group {groupId}", 2, 0);
            RefreshGroupLists();
        }

        private void AddGroupRow(int groupId, string name, int colorIndex, int lineStyleIndex)
        {
            var rowGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var buttonPanel = new StackPanel
            {
                Margin = new Thickness(0, 36, 8, 0),
                VerticalAlignment = VerticalAlignment.Top
            };

            var addButton = new Button
            {
                Content = ">",
                Width = 28,
                Height = 28,
                Padding = new Thickness(0),
                Tag = groupId
            };
            addButton.Click += AddToGroupButton_Click;

            var removeButton = new Button
            {
                Content = "<",
                Width = 28,
                Height = 28,
                Padding = new Thickness(0),
                Margin = new Thickness(0, 6, 0, 0),
                Tag = groupId
            };
            removeButton.Click += RemoveFromGroupButton_Click;

            buttonPanel.Children.Add(addButton);
            buttonPanel.Children.Add(removeButton);
            Grid.SetColumn(buttonPanel, 0);
            rowGrid.Children.Add(buttonPanel);

            var groupBorder = new Border
            {
                BorderBrush = System.Windows.Media.Brushes.LightGray,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(8)
            };

            var groupGrid = new Grid();
            groupGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            groupGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(170) });
            groupGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            groupGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            groupGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });

            var nameTextBox = new TextBox
            {
                Text = string.IsNullOrWhiteSpace(name) ? $"Group {groupId}" : name,
                Margin = new Thickness(0, 0, 6, 6)
            };
            Grid.SetRow(nameTextBox, 0);
            Grid.SetColumn(nameTextBox, 0);
            groupGrid.Children.Add(nameTextBox);

            var colorCombo = new ComboBox { Margin = new Thickness(0, 0, 6, 6) };
            BindColorCombo(colorCombo, colorIndex);
            Grid.SetRow(colorCombo, 0);
            Grid.SetColumn(colorCombo, 1);
            groupGrid.Children.Add(colorCombo);

            var lineStyleCombo = new ComboBox { Margin = new Thickness(0, 0, 0, 6) };
            BindLineStyleCombo(lineStyleCombo, lineStyleIndex);
            Grid.SetRow(lineStyleCombo, 0);
            Grid.SetColumn(lineStyleCombo, 2);
            groupGrid.Children.Add(lineStyleCombo);

            var columnsList = new ListBox { SelectionMode = SelectionMode.Extended };
            Grid.SetRow(columnsList, 1);
            Grid.SetColumn(columnsList, 0);
            Grid.SetColumnSpan(columnsList, 3);
            groupGrid.Children.Add(columnsList);

            groupBorder.Child = groupGrid;
            Grid.SetColumn(groupBorder, 1);
            rowGrid.Children.Add(groupBorder);

            GroupRowsHost.Children.Add(rowGrid);
            _groupRows.Add(new GroupRow
            {
                GroupId = groupId,
                Container = groupBorder,
                NameTextBox = nameTextBox,
                ColorComboBox = colorCombo,
                LineStyleComboBox = lineStyleCombo,
                GroupColumnsListBox = columnsList
            });
        }

        private void BindColorCombo(ComboBox comboBox, int selectedIndex)
        {
            comboBox.ItemsSource = _colorOptions;
            comboBox.DisplayMemberPath = "Name";
            comboBox.SelectedIndex = Math.Max(0, Math.Min(selectedIndex, _colorOptions.Count - 1));
        }

        private void BindLineStyleCombo(ComboBox comboBox, int selectedIndex)
        {
            comboBox.ItemsSource = _lineStyleOptions;
            comboBox.SelectedIndex = Math.Max(0, Math.Min(selectedIndex, _lineStyleOptions.Count - 1));
        }

        private void ReloadColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            ReloadColumns();
        }

        private void ReloadColumns()
        {
            try
            {
                if (!File.Exists(_fileInfo.FilePath))
                {
                    MessageBox.Show("File not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!int.TryParse(HeaderRowTextBox.Text, out int headerRowNumber) || headerRowNumber <= 0)
                {
                    MessageBox.Show("Header Row must be a positive number.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var lines = File.ReadAllLines(_fileInfo.FilePath);
                if (lines.Length == 0 || headerRowNumber > lines.Length)
                {
                    MessageBox.Show("Header Row is out of range.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string delimiter = GetSelectedDelimiter();
                _headers = GraphMakerTableHelper.SplitLine(lines[headerRowNumber - 1], delimiter).ToList();

                if (_headers.Count == 0)
                {
                    MessageBox.Show("No header columns detected.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _columnGroups.Keys.Where(k => k >= _headers.Count).ToList().ForEach(k => _columnGroups.Remove(k));
                PopulateControls();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to parse file.\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PopulateControls()
        {
            PopulateAxisControls();
            PopulateLimitControls();
            RefreshGroupLists();
            UpdateXAxisModeUI();
        }

        private void PopulateAxisControls()
        {
            XAxisComboBox.Items.Clear();
            XAxisComboBox.Items.Add("(None)");
            foreach (string header in _headers)
            {
                XAxisComboBox.Items.Add(header);
            }

            XAxisComboBox.SelectedIndex = GetSafeIndex(_fileInfo.SavedXAxisIndex, XAxisComboBox.Items.Count, _headers.Count > 0 ? 1 : 0);
        }

        private void PopulateLimitControls()
        {
            RefreshLimitLists();
        }

        private void RefreshLimitLists()
        {
            int refIndex = _fileInfo.SavedSpecColumnIndex > 0 ? _fileInfo.SavedSpecColumnIndex - 1 : -1;
            int upperIndex = _fileInfo.SavedUpperLimitColumnIndex > 0 ? _fileInfo.SavedUpperLimitColumnIndex - 1 : -1;
            int lowerIndex = _fileInfo.SavedLowerLimitColumnIndex > 0 ? _fileInfo.SavedLowerLimitColumnIndex - 1 : -1;

            SetLimitListItem(RefLimitListBox, refIndex);
            SetLimitListItem(UpperLimitListBox, upperIndex);
            SetLimitListItem(LowerLimitListBox, lowerIndex);
        }

        private void SetLimitListItem(ListBox target, int index)
        {
            target.Items.Clear();
            if (index >= 0 && index < _headers.Count)
            {
                target.Items.Add(new ColumnItem { Index = index, Header = _headers[index] });
            }
        }

        private void MoveSelectedToLimit(ListBox target)
        {
            var selected = UnassignedColumnsListBox.SelectedItems.OfType<ColumnItem>().FirstOrDefault();
            if (selected == null)
            {
                return;
            }

            target.Items.Clear();
            target.Items.Add(new ColumnItem { Index = selected.Index, Header = selected.Header });
        }

        private void RemoveFromLimit(ListBox target)
        {
            target.Items.Clear();
        }

        private void MoveToRefButton_Click(object sender, RoutedEventArgs e) => MoveSelectedToLimit(RefLimitListBox);
        private void RemoveFromRefButton_Click(object sender, RoutedEventArgs e) => RemoveFromLimit(RefLimitListBox);
        private void MoveToUpperButton_Click(object sender, RoutedEventArgs e) => MoveSelectedToLimit(UpperLimitListBox);
        private void RemoveFromUpperButton_Click(object sender, RoutedEventArgs e) => RemoveFromLimit(UpperLimitListBox);
        private void MoveToLowerButton_Click(object sender, RoutedEventArgs e) => MoveSelectedToLimit(LowerLimitListBox);
        private void RemoveFromLowerButton_Click(object sender, RoutedEventArgs e) => RemoveFromLimit(LowerLimitListBox);

        private void RefreshGroupLists()
        {
            var validGroupIds = _groupRows.Select(g => g.GroupId).ToHashSet();
            var invalidColumns = _columnGroups.Where(kv => !validGroupIds.Contains(kv.Value)).Select(kv => kv.Key).ToList();
            foreach (int col in invalidColumns)
            {
                _columnGroups.Remove(col);
            }

            UnassignedColumnsListBox.Items.Clear();
            for (int i = 0; i < _headers.Count; i++)
            {
                if (!_columnGroups.ContainsKey(i))
                {
                    UnassignedColumnsListBox.Items.Add(new ColumnItem { Index = i, Header = _headers[i] });
                }
            }

            foreach (var row in _groupRows)
            {
                row.GroupColumnsListBox.Items.Clear();
                for (int i = 0; i < _headers.Count; i++)
                {
                    if (_columnGroups.TryGetValue(i, out int groupId) && groupId == row.GroupId)
                    {
                        row.GroupColumnsListBox.Items.Add(new ColumnItem { Index = i, Header = _headers[i] });
                    }
                }
            }
        }

        private void AddToGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button button || button.Tag is not int groupId)
            {
                return;
            }

            if (UnassignedColumnsListBox.SelectedItems.Count == 0)
            {
                MessageBox.Show("Select one or more headers on the left list.", "Group Assign", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            foreach (var selected in UnassignedColumnsListBox.SelectedItems.OfType<ColumnItem>())
            {
                _columnGroups[selected.Index] = groupId;
            }

            RefreshGroupLists();
        }

        private void RemoveFromGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button button || button.Tag is not int groupId)
            {
                return;
            }

            var row = _groupRows.FirstOrDefault(g => g.GroupId == groupId);
            if (row == null)
            {
                return;
            }

            foreach (var selected in row.GroupColumnsListBox.SelectedItems.OfType<ColumnItem>().ToList())
            {
                _columnGroups.Remove(selected.Index);
            }

            RefreshGroupLists();
        }

        private static int GetSafeIndex(int value, int count, int fallback)
        {
            if (count <= 0)
            {
                return -1;
            }

            if (value >= 0 && value < count)
            {
                return value;
            }

            return Math.Max(0, Math.Min(fallback, count - 1));
        }

        private void XAxisModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateXAxisModeUI();
        }

        private void UpdateXAxisModeUI()
        {
            bool useColumn = XAxisModeComboBox.SelectedIndex == 0;
            XAxisComboBox.IsEnabled = useColumn;
            if (!useColumn)
            {
                XAxisComboBox.SelectedIndex = 0;
            }
            else if (XAxisComboBox.SelectedIndex <= 0 && XAxisComboBox.Items.Count > 1)
            {
                XAxisComboBox.SelectedIndex = 1;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out int headerRowNumber) || headerRowNumber <= 0)
            {
                MessageBox.Show("Header Row must be a positive number.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedYIndices = _columnGroups.Keys
                .Where(i => i >= 0 && i < _headers.Count)
                .Distinct()
                .OrderBy(i => i)
                .ToList();

            if (selectedYIndices.Count == 0)
            {
                MessageBox.Show("Please select at least one Y column.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _fileInfo.Delimiter = GetSelectedDelimiter();
            _fileInfo.HeaderRowNumber = headerRowNumber;
            _fileInfo.SavedXAxisIndex = XAxisModeComboBox.SelectedIndex == 0 ? Math.Max(0, XAxisComboBox.SelectedIndex) : 0;
            _fileInfo.SavedYAxisIndices = selectedYIndices;
            _fileInfo.SavedXAxisLogScale = XAxisLogScaleCheckBox.IsChecked == true;

            _fileInfo.SavedSpecColumnIndex = RefLimitListBox.Items.Count > 0 && RefLimitListBox.Items[0] is ColumnItem refItem ? refItem.Index + 1 : 0;
            _fileInfo.SavedUpperLimitColumnIndex = UpperLimitListBox.Items.Count > 0 && UpperLimitListBox.Items[0] is ColumnItem upperItem ? upperItem.Index + 1 : 0;
            _fileInfo.SavedLowerLimitColumnIndex = LowerLimitListBox.Items.Count > 0 && LowerLimitListBox.Items[0] is ColumnItem lowerItem ? lowerItem.Index + 1 : 0;
            _fileInfo.SavedSpecColorIndex = SpecColorComboBox.SelectedIndex;
            _fileInfo.SavedUpperColorIndex = UpperLimitColorComboBox.SelectedIndex;
            _fileInfo.SavedLowerColorIndex = LowerLimitColorComboBox.SelectedIndex;
            _fileInfo.SavedSpecLineStyleIndex = SpecLineStyleComboBox.SelectedIndex;
            _fileInfo.SavedUpperLineStyleIndex = UpperLimitLineStyleComboBox.SelectedIndex;
            _fileInfo.SavedLowerLineStyleIndex = LowerLimitLineStyleComboBox.SelectedIndex;

            var validGroupIds = _groupRows.Select(g => g.GroupId).ToHashSet();
            var filteredColumnMap = _columnGroups
                .Where(kv => validGroupIds.Contains(kv.Value))
                .ToDictionary(kv => kv.Key, kv => kv.Value);
            _fileInfo.SavedColumnGroups = filteredColumnMap;

            _fileInfo.SavedGroups = _groupRows
                .Select(g => new SavedGroupSetting
                {
                    GroupId = g.GroupId,
                    Name = string.IsNullOrWhiteSpace(g.NameTextBox.Text) ? $"Group {g.GroupId}" : g.NameTextBox.Text.Trim(),
                    ColorIndex = g.ColorComboBox.SelectedIndex < 0 ? 0 : g.ColorComboBox.SelectedIndex,
                    LineStyleIndex = g.LineStyleComboBox.SelectedIndex < 0 ? 0 : g.LineStyleComboBox.SelectedIndex,
                    RefColumnIndex = -1,
                    UpperLimitColumnIndex = -1,
                    LowerLimitColumnIndex = -1
                })
                .OrderBy(g => g.GroupId)
                .ToList();

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private string GetSelectedDelimiter()
        {
            if (CommaDelimiterRadio.IsChecked == true) return ",";
            if (SpaceDelimiterRadio.IsChecked == true) return " ";
            return "\t";
        }

    }
}
