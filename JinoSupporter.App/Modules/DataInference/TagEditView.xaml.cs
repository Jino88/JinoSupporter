using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace JinoSupporter.App.Modules.DataInference;

public partial class TagEditView : UserControl
{
    private readonly DataInferenceRepository _repository = new();
    private string? _selectedTag;

    public TagEditView()
    {
        InitializeComponent();
        Loaded += (_, _) => LoadTags();
    }

    private void LoadTags()
    {
        TagListPanel.Children.Clear();
        _selectedTag = null;
        CurrentTagBox.Text = string.Empty;
        NewTagBox.Text     = string.Empty;
        RenameButton.IsEnabled = false;
        AffectedLabel.Text = string.Empty;
        ResultBorder.Visibility = Visibility.Collapsed;

        List<string> tags = _repository.GetAllDistinctTags();

        if (tags.Count == 0)
        {
            TagListPanel.Children.Add(new TextBlock
            {
                Text       = "No registered tags.",
                FontSize   = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                Margin     = new Thickness(4, 6, 4, 6)
            });
            StatusTextBlock.Text = "No tags.";
            return;
        }

        foreach (string tag in tags)
        {
            string captured = tag;
            var btn = new Button
            {
                Content             = tag,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Padding             = new Thickness(8, 5, 8, 5),
                Margin              = new Thickness(0, 0, 0, 2),
                FontSize            = 12,
                Background          = Brushes.Transparent,
                BorderThickness     = new Thickness(0),
                Foreground          = new SolidColorBrush(Color.FromRgb(34, 48, 74)),
                Cursor              = Cursors.Hand
            };
            btn.Click += (_, _) => SelectTag(captured);
            TagListPanel.Children.Add(btn);
        }

        StatusTextBlock.Text = $"{tags.Count} tag(s)";
        SelectedTagLabel.Text = "← Select a tag from the list on the left";
    }

    private void SelectTag(string tag)
    {
        _selectedTag = tag;
        CurrentTagBox.Text = tag;
        NewTagBox.Text     = tag;
        NewTagBox.SelectAll();
        NewTagBox.Focus();
        RenameButton.IsEnabled = true;
        ResultBorder.Visibility = Visibility.Collapsed;
        SelectedTagLabel.Text = $"Selected tag: {tag}";
        AffectedLabel.Text    = string.Empty;

        // Background highlight
        foreach (Button b in TagListPanel.Children.OfType<Button>())
        {
            bool isSelected = string.Equals(b.Content as string, tag, StringComparison.OrdinalIgnoreCase);
            b.Background = isSelected
                ? new SolidColorBrush(Color.FromRgb(219, 234, 254))
                : Brushes.Transparent;
        }
    }

    private void RenameTag()
    {
        if (_selectedTag is null) return;

        string newName = NewTagBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(newName))
        {
            StatusTextBlock.Text = "Enter the new tag name.";
            return;
        }

        if (string.Equals(newName, _selectedTag, StringComparison.Ordinal))
        {
            StatusTextBlock.Text = "The name is the same.";
            return;
        }

        int affected = _repository.RenameTag(_selectedTag, newName);

        ResultBorder.Visibility = Visibility.Visible;
        ResultBorder.Background  = new SolidColorBrush(Color.FromRgb(236, 253, 245));
        ResultBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(110, 231, 183));
        ResultTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(6, 95, 70));
        ResultTextBlock.Text = $"'{_selectedTag}' → '{newName}' renamed ({affected} dataset(s) updated)";

        StatusTextBlock.Text = $"Tag renamed: {affected} dataset(s) updated.";
        LoadTags();
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e) => LoadTags();

    private void RenameButton_Click(object sender, RoutedEventArgs e) => RenameTag();

    private void NewTagBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            RenameTag();
            e.Handled = true;
        }
    }
}
