using DataMaker.R6.FetchDataBMES;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using VideoConverter;
using WorkbenchHost.Infrastructure;
using WorkbenchHost.Modules.ScreenCapture;

namespace WorkbenchHost.Modules.AppSettings;

public partial class AppSettingsView : UserControl
{
    private sealed class SettingsEntry
    {
        public string Name { get; init; } = string.Empty;
        public string Description { get; init; } = string.Empty;
        public string Path { get; init; } = string.Empty;
    }

    private readonly ObservableCollection<SettingsEntry> _entries = [];

    public AppSettingsView()
    {
        InitializeComponent();
        SettingsItemsControl.ItemsSource = _entries;
        RefreshFromCurrentSettings();
    }

    private void BrowseBatchDirectoryButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select a folder to use for WorkHost settings and data",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            SelectedPath = BatchDirectoryTextBox.Text.Trim()
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            BatchDirectoryTextBox.Text = dialog.SelectedPath;
        }
    }

    private void ApplyBatchDirectoryButton_Click(object sender, RoutedEventArgs e)
    {
        string directory = BatchDirectoryTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(directory))
        {
            MessageBox.Show("Select a batch directory first.", "Settings", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            string normalizedDirectory = Path.GetFullPath(directory);
            Directory.CreateDirectory(normalizedDirectory);

            WorkbenchSettingsStore.SetSettingsFilePath(Path.Combine(normalizedDirectory, "workhost-settings.json"));
            AppSettingsPathManager.SetStorageRootDirectory(normalizedDirectory);
            RefreshFromCurrentSettings();
            StatusTextBlock.Text = "Batch directory applied to settings file and shared storage root.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to apply batch directory.\n{ex.Message}", "Settings", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenBatchDirectoryButton_Click(object sender, RoutedEventArgs e)
    {
        OpenPath(BatchDirectoryTextBox.Text.Trim());
    }

    private void BrowseSettingsFileButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "Select unified settings file",
            Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
            FileName = Path.GetFileName(SettingsFilePathTextBox.Text),
            InitialDirectory = ResolveInitialDirectory(SettingsFilePathTextBox.Text)
        };

        if (dialog.ShowDialog() == true)
        {
            SettingsFilePathTextBox.Text = dialog.FileName;
        }
    }

    private void SaveSettingsFileButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            WorkbenchSettingsStore.SetSettingsFilePath(SettingsFilePathTextBox.Text.Trim());
            RefreshFromCurrentSettings();
            StatusTextBlock.Text = "Unified settings file path saved.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save settings file path.\n{ex.Message}", "Settings", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ResetSettingsFileButton_Click(object sender, RoutedEventArgs e)
    {
        SettingsFilePathTextBox.Text = WorkbenchSettingsStore.DefaultSettingsFilePath;
        StatusTextBlock.Text = "Settings file path reset to default. Save File Path to apply.";
    }

    private void OpenSettingsFileFolderButton_Click(object sender, RoutedEventArgs e)
    {
        OpenPath(SettingsFilePathTextBox.Text.Trim());
    }

    private void BrowseStorageRootButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select shared settings root folder",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            SelectedPath = StorageRootTextBox.Text.Trim()
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            StorageRootTextBox.Text = dialog.SelectedPath;
        }
    }

    private void SaveStorageRootButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AppSettingsPathManager.SetStorageRootDirectory(StorageRootTextBox.Text.Trim());
            RefreshFromCurrentSettings();
            StatusTextBlock.Text = "Shared storage root saved.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save storage root.\n{ex.Message}", "Settings", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ResetStorageRootButton_Click(object sender, RoutedEventArgs e)
    {
        StorageRootTextBox.Text = AppSettingsPathManager.DefaultStorageRootDirectory;
        StatusTextBlock.Text = "Storage root reset to default path. Save Root to apply.";
    }

    private void OpenStorageRootButton_Click(object sender, RoutedEventArgs e)
    {
        OpenPath(StorageRootTextBox.Text.Trim());
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshFromCurrentSettings();
        StatusTextBlock.Text = "Settings list refreshed.";
    }

    private void OpenEntryButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string path)
        {
            OpenPath(path);
        }
    }

    private void RefreshFromCurrentSettings()
    {
        WorkbenchSettingsStore.WorkbenchSettings settings = WorkbenchSettingsStore.GetSettings();
        BatchDirectoryTextBox.Text = settings.StorageRootDirectory;
        SettingsFilePathTextBox.Text = WorkbenchSettingsStore.SettingsFilePath;
        StorageRootTextBox.Text = AppSettingsPathManager.StorageRootDirectory;

        _entries.Clear();
        AddEntry("Bootstrap pointer", "Custom settings file path pointer", WorkbenchSettingsStore.BootstrapFilePath);
        AddEntry("Unified settings JSON", "All WorkHost settings values are saved here", WorkbenchSettingsStore.SettingsFilePath);
        AddEntry("Storage root", "Default root for module data/output files", settings.StorageRootDirectory);
        AddEntry("ScreenCapture hotkeys", "Stored inside the unified settings JSON", CaptureHotkeyManager.SettingsPath);
        AddEntry("VideoConverter settings", "Stored inside the unified settings JSON", VideoConverter.MainWindow.SettingsFilePath);
        AddEntry("Memo database path", "Current memo DB target", settings.Memo.DatabasePath);
        AddEntry("DiskTree database", "Indexed file metadata database", DiskTree.MainWindow.DefaultDatabasePath);
        AddEntry("BMES credentials", "Stored inside the unified settings JSON", FormSettingBMESWindow.GetCurrentDataFilePath());
    }

    private void AddEntry(string name, string description, string path)
    {
        _entries.Add(new SettingsEntry
        {
            Name = name,
            Description = description,
            Path = path
        });
    }

    private static string ResolveInitialDirectory(string path)
    {
        string? directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
        {
            return directory;
        }

        return AppSettingsPathManager.StorageRootDirectory;
    }

    private static void OpenPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        string normalizedPath = Path.GetFullPath(path);
        string? directory = Directory.Exists(normalizedPath)
            ? normalizedPath
            : Path.GetDirectoryName(normalizedPath);

        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (File.Exists(normalizedPath))
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{normalizedPath}\"") { UseShellExecute = true });
            return;
        }

        if (Directory.Exists(normalizedPath))
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{normalizedPath}\"") { UseShellExecute = true });
            return;
        }

        if (!string.IsNullOrWhiteSpace(directory))
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{directory}\"") { UseShellExecute = true });
        }
    }
}
