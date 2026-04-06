using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace GraphMaker
{
    public sealed class ProcessPairGroupOption : INotifyPropertyChanged
    {
        public int GroupId { get; init; }
        private string _name = string.Empty;
        public string Name
        {
            get => _name;
            set
            {
                if (_name == value)
                {
                    return;
                }
                _name = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public sealed class ProcessPairSelectionItem
    {
        public int ProcessIndex { get; init; }
        public string DisplayName { get; init; } = string.Empty;
        public bool IsSelected { get; set; }
        public int GroupId { get; set; } = 1;
    }

    public partial class ProcessPairSelectionWindow : Window
    {
        private readonly ObservableCollection<ProcessPairSelectionItem> _items = new();
        private readonly ObservableCollection<ProcessPairGroupOption> _groupOptions = new();

        public IReadOnlyList<ProcessPairSelectionItem> SelectedItems =>
            _items.Where(i => i.IsSelected).ToList();
        public IReadOnlyList<ProcessPairGroupOption> GroupOptions => _groupOptions;

        public ProcessPairSelectionWindow(
            IReadOnlyList<string> processNames,
            ISet<int> selectedProcessIndices,
            IReadOnlyDictionary<int, int> processGroupMap,
            IReadOnlyDictionary<int, string> groupNameMap)
        {
            InitializeComponent();
            DataContext = this;

            if (groupNameMap.Count > 0)
            {
                foreach (var kv in groupNameMap.OrderBy(kv => kv.Key))
                {
                    if (kv.Key <= 0)
                    {
                        continue;
                    }

                    _groupOptions.Add(new ProcessPairGroupOption
                    {
                        GroupId = kv.Key,
                        Name = string.IsNullOrWhiteSpace(kv.Value) ? $"Group {kv.Key}" : kv.Value
                    });
                }
            }

            if (_groupOptions.Count == 0)
            {
                _groupOptions.Add(new ProcessPairGroupOption { GroupId = 1, Name = "Group 1" });
                _groupOptions.Add(new ProcessPairGroupOption { GroupId = 2, Name = "Group 2" });
            }

            int defaultGroup = _groupOptions[0].GroupId;

            for (int processIndex = 0; processIndex < processNames.Count; processIndex++)
            {
                int assignedGroup = processGroupMap.TryGetValue(processIndex, out int g) && _groupOptions.Any(o => o.GroupId == g)
                    ? g
                    : defaultGroup;

                _items.Add(new ProcessPairSelectionItem
                {
                    ProcessIndex = processIndex,
                    DisplayName = processNames[processIndex],
                    IsSelected = selectedProcessIndices.Count == 0 || selectedProcessIndices.Contains(processIndex),
                    GroupId = assignedGroup
                });
            }

            PairListBox.ItemsSource = _items;
            if (_groupOptions.Count > 0)
            {
                GroupManageComboBox.SelectedIndex = 0;
                GroupNameTextBox.Text = _groupOptions[0].Name;
            }
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items)
            {
                item.IsSelected = true;
            }

            PairListBox.Items.Refresh();
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items)
            {
                item.IsSelected = false;
            }

            PairListBox.Items.Refresh();
        }

        private void AddGroupButton_Click(object sender, RoutedEventArgs e)
        {
            int nextId = _groupOptions.Count == 0 ? 1 : _groupOptions.Max(g => g.GroupId) + 1;
            _groupOptions.Add(new ProcessPairGroupOption
            {
                GroupId = nextId,
                Name = $"Group {nextId}"
            });
            GroupManageComboBox.SelectedItem = _groupOptions.Last();
            PairListBox.Items.Refresh();
        }

        private void RemoveGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (_groupOptions.Count <= 1)
            {
                MessageBox.Show("At least one group must remain.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var groupToRemove = _groupOptions.Last();
            _groupOptions.Remove(groupToRemove);
            int fallbackGroupId = _groupOptions[0].GroupId;

            foreach (var item in _items.Where(i => i.GroupId == groupToRemove.GroupId))
            {
                item.GroupId = fallbackGroupId;
            }

            GroupManageComboBox.SelectedIndex = 0;
            GroupNameTextBox.Text = _groupOptions[0].Name;
            PairListBox.Items.Refresh();
        }

        private void GroupManageComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (GroupManageComboBox.SelectedItem is ProcessPairGroupOption option)
            {
                GroupNameTextBox.Text = option.Name;
            }
        }

        private void ApplyGroupNameButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupManageComboBox.SelectedItem is not ProcessPairGroupOption option)
            {
                return;
            }

            var newName = (GroupNameTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(newName))
            {
                MessageBox.Show("Group name cannot be empty.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            option.Name = newName;
            PairListBox.Items.Refresh();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
