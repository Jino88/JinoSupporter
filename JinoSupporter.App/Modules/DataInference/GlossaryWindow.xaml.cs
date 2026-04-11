using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.App.Modules.DataInference;

public partial class GlossaryWindow : Window
{
    private readonly DataInferenceRepository _repository;
    private readonly ObservableCollection<GlossaryEntry> _entries = [];

    public GlossaryWindow(DataInferenceRepository repository)
    {
        InitializeComponent();
        _repository = repository;
        GlossaryGrid.ItemsSource = _entries;
        LoadFromRepository();
    }

    private void LoadFromRepository()
    {
        _entries.Clear();
        foreach ((string term, string desc) in _repository.GetGlossaryEntriesNewestFirst())
            _entries.Add(new GlossaryEntry { Term = term, Description = desc });
        StatusText.Text = $"{_entries.Count} entry(ies)";
    }

    private void AddRowButton_Click(object sender, RoutedEventArgs e)
    {
        var entry = new GlossaryEntry();
        _entries.Add(entry);
        GlossaryGrid.ScrollIntoView(entry);
        GlossaryGrid.SelectedItem = entry;
        GlossaryGrid.CurrentCell  = new DataGridCellInfo(entry, GlossaryGrid.Columns[0]);
        GlossaryGrid.BeginEdit();
    }

    private void DeleteRowButton_Click(object sender, RoutedEventArgs e)
    {
        if (GlossaryGrid.SelectedItem is GlossaryEntry entry)
            _entries.Remove(entry);
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        GlossaryGrid.CommitEdit(DataGridEditingUnit.Row, exitEditingMode: true);

        _repository.ReplaceGlossary(
            _entries
                .Where(en => !string.IsNullOrWhiteSpace(en.Term))
                .Select(en => (en.Term, en.Description)));

        // Re-read to pick up sort / dedup
        LoadFromRepository();
        StatusText.Text = $"Saved — {_entries.Count} entry(ies)";
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
}

public sealed class GlossaryEntry : INotifyPropertyChanged
{
    private string _term        = string.Empty;
    private string _description = string.Empty;

    public string Term
    {
        get => _term;
        set { _term = value; OnPropertyChanged(); }
    }

    public string Description
    {
        get => _description;
        set { _description = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
