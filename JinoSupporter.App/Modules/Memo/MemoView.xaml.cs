using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using WorkbenchHost.Infrastructure;

namespace WorkbenchHost.Modules.Memo;

public partial class MemoView : UserControl
{
    private sealed class MemoImageItem
    {
        public long Id { get; set; }
        public required byte[] ImageBytes { get; init; }
        public required BitmapImage Preview { get; init; }
    }

    private sealed class MemoRow : INotifyPropertyChanged
    {
        private long _id;
        private string _memoDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        private string _model = string.Empty;
        private string _category = string.Empty;
        private string _tags = string.Empty;
        private string _issue = string.Empty;
        private readonly ObservableCollection<MemoImageItem> _images = new();
        private readonly List<long> _deletedImageIds = [];
        private bool _isDirty = true;

        public long Id
        {
            get => _id;
            set
            {
                if (_id == value)
                {
                    return;
                }

                _id = value;
                OnPropertyChanged(nameof(Id));
            }
        }

        public string MemoDate
        {
            get => _memoDate;
            set
            {
                if (_memoDate == value)
                {
                    return;
                }

                _memoDate = value;
                _isDirty = true;
                OnPropertyChanged(nameof(MemoDate));
            }
        }

        public string Model
        {
            get => _model;
            set
            {
                if (_model == value)
                {
                    return;
                }

                _model = value;
                _isDirty = true;
                OnPropertyChanged(nameof(Model));
            }
        }

        public string Category
        {
            get => _category;
            set
            {
                if (_category == value)
                {
                    return;
                }

                _category = value;
                _isDirty = true;
                OnPropertyChanged(nameof(Category));
            }
        }

        public string Issue
        {
            get => _issue;
            set
            {
                if (_issue == value)
                {
                    return;
                }

                _issue = value;
                _isDirty = true;
                OnPropertyChanged(nameof(Issue));
            }
        }

        public string Tags
        {
            get => _tags;
            set
            {
                if (_tags == value)
                {
                    return;
                }

                _tags = value;
                _isDirty = true;
                OnPropertyChanged(nameof(Tags));
                OnPropertyChanged(nameof(TagItems));
            }
        }

        public ObservableCollection<MemoImageItem> Images => _images;
        public IReadOnlyList<string> TagItems => ParseTags(_tags);
        public bool IsDirty => _isDirty;
        public IReadOnlyList<long> DeletedImageIds => _deletedImageIds;

        public event PropertyChangedEventHandler? PropertyChanged;

        public void AddImage(byte[] imageBytes)
        {
            _images.Add(new MemoImageItem
            {
                Id = 0,
                ImageBytes = imageBytes,
                Preview = CreatePreviewImage(imageBytes)
            });
            _isDirty = true;
            OnPropertyChanged(nameof(Images));
        }

        public void AddLoadedImage(long id, byte[] imageBytes)
        {
            _images.Add(new MemoImageItem
            {
                Id = id,
                ImageBytes = imageBytes,
                Preview = CreatePreviewImage(imageBytes)
            });
        }

        public void RemoveImage(MemoImageItem image)
        {
            if (_images.Remove(image))
            {
                if (image.Id > 0)
                {
                    _deletedImageIds.Add(image.Id);
                }

                _isDirty = true;
                OnPropertyChanged(nameof(Images));
            }
        }

        public void MarkClean()
        {
            _isDirty = false;
            _deletedImageIds.Clear();
        }

        private static BitmapImage CreatePreviewImage(byte[] imageBytes)
        {
            using MemoryStream stream = new(imageBytes);
            BitmapImage image = new();
            image.BeginInit();
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.StreamSource = stream;
            image.EndInit();
            image.Freeze();
            return image;
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static IReadOnlyList<string> ParseTags(string tags)
        {
            return tags
                .Split([',', ';', '\r', '\n', ' '], StringSplitOptions.RemoveEmptyEntries)
                .Select(static tag => tag.StartsWith('#') ? tag : $"#{tag}")
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }

    private readonly ObservableCollection<MemoRow> _rows = new();
    private readonly HashSet<long> _deletedRowIds = [];
    private MemoRow? _selectedRow;

    public event Action? WebModuleSnapshotChanged;

    public MemoView()
    {
        InitializeComponent();
        MemoDataGrid.ItemsSource = _rows;
        LoadRowsFromCurrentDatabase();
    }

    private void AddRowButton_Click(object sender, RoutedEventArgs e)
    {
        AddNewRow();
        StatusTextBlock.Text = $"Row added. Total rows: {_rows.Count:N0}.";
        NotifyWebModuleSnapshotChanged();
    }

    private void DeleteRowButton_Click(object sender, RoutedEventArgs e)
    {
        MemoRow? selectedRow = MemoDataGrid.SelectedItem as MemoRow;
        if (selectedRow is null && MemoDataGrid.CurrentItem is MemoRow currentRow)
        {
            selectedRow = currentRow;
        }

        if (selectedRow is null)
        {
            MessageBox.Show("Select a row first.", "Memo",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _rows.Remove(selectedRow);
        if (selectedRow.Id > 0)
        {
            _deletedRowIds.Add(selectedRow.Id);
        }

        if (_rows.Count == 0)
        {
            AddNewRow();
        }

        StatusTextBlock.Text = $"Row deleted. Total rows: {_rows.Count:N0}.";
        NotifyWebModuleSnapshotChanged();
    }

    private void SelectRowButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.DataContext is not MemoRow row)
        {
            return;
        }

        MemoDataGrid.SelectedItems.Clear();
        MemoDataGrid.SelectedItem = row;
        MemoDataGrid.CurrentItem = row;
        MemoDataGrid.ScrollIntoView(row);
        _selectedRow = row;
        StatusTextBlock.Text = "Memo row selected.";
        NotifyWebModuleSnapshotChanged();
    }

    private void PasteImageButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.DataContext is not MemoRow row)
        {
            return;
        }

        try
        {
            int imageCount = AddClipboardImages(row);
            if (imageCount == 0)
            {
                MessageBox.Show("Clipboard does not contain an image or image file list.", "Memo",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            StatusTextBlock.Text = $"Added {imageCount:N0} image(s) to the selected memo row.";
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save clipboard image: {ex.Message}", "Memo",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ImagePreviewButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.DataContext is not MemoImageItem imageItem)
        {
            return;
        }

        var viewer = new MemoImageViewerWindow(imageItem.Preview)
        {
            Owner = Window.GetWindow(this)
        };
        viewer.ShowDialog();
    }

    private void DeleteImageMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem menuItem ||
            menuItem.DataContext is not MemoImageItem imageItem ||
            FindOwningMemoRow(imageItem) is not MemoRow row)
        {
            return;
        }

        row.RemoveImage(imageItem);
        StatusTextBlock.Text = "Selected image removed from memo row.";
        NotifyWebModuleSnapshotChanged();
    }

    private void TagButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Content is not string tag || string.IsNullOrWhiteSpace(tag))
        {
            return;
        }

        SortRowsByTag(tag);
        StatusTextBlock.Text = $"Sorted memo rows by tag {tag}.";
        NotifyWebModuleSnapshotChanged();
    }

    private void MemoEditingTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || Keyboard.Modifiers != ModifierKeys.Alt || sender is not TextBox textBox)
        {
            return;
        }

        int caretIndex = textBox.CaretIndex;
        textBox.Text = textBox.Text.Insert(caretIndex, Environment.NewLine);
        textBox.CaretIndex = caretIndex + Environment.NewLine.Length;
        e.Handled = true;
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string databasePath = GetDatabasePath();
            SaveRowsToDatabase(databasePath);
            StatusTextBlock.Text = $"Saved {_rows.Count:N0} memo row(s) to DB.";
            MessageBox.Show("Memo data saved successfully.", "Memo",
                MessageBoxButton.OK, MessageBoxImage.Information);
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save memo data: {ex.Message}", "Memo",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void AddNewRow()
    {
        MemoRow row = new();
        _rows.Add(row);
        _selectedRow = row;
    }

    private string GetDatabasePath()
    {
        string path = LoadSavedDatabasePath();
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("Enter a DB path first.");
        }

        string? directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return path;
    }

    private static string LoadSavedDatabasePath()
    {
        try
        {
            string databasePath = WorkbenchSettingsStore.GetSettings().Memo.DatabasePath;
            if (!string.IsNullOrWhiteSpace(databasePath))
            {
                return databasePath;
            }
        }
        catch
        {
        }

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "JinoWorkHost",
            "memo.db");
    }

    private int AddClipboardImages(MemoRow row)
    {
        int addedCount = 0;

        if (Clipboard.ContainsFileDropList())
        {
            var files = Clipboard.GetFileDropList();
            foreach (string file in files)
            {
                if (string.IsNullOrWhiteSpace(file) || !IsImageFile(file))
                {
                    continue;
                }

                row.AddImage(File.ReadAllBytes(file));
                addedCount++;
            }

            return addedCount;
        }

        if (!Clipboard.ContainsImage())
        {
            return addedCount;
        }

        BitmapSource image = Clipboard.GetImage();
        using MemoryStream stream = new();
        BitmapEncoder encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(image));
        encoder.Save(stream);
        row.AddImage(stream.ToArray());
        return addedCount + 1;
    }

    private static bool IsImageFile(string path)
    {
        string extension = Path.GetExtension(path);
        return extension.Equals(".png", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".bmp", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".gif", StringComparison.OrdinalIgnoreCase);
    }

    private void SaveRowsToDatabase(string databasePath)
    {
        using SqliteConnection connection = new($"Data Source={databasePath}");
        connection.Open();

        using SqliteCommand createMemoTable = connection.CreateCommand();
        createMemoTable.CommandText = """
            CREATE TABLE IF NOT EXISTS MemoEntries (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                MemoDate TEXT NOT NULL,
                Model TEXT,
                Category TEXT,
                Tags TEXT,
                Issue TEXT,
                CreatedAtUtc TEXT NOT NULL
            );
            """;
        createMemoTable.ExecuteNonQuery();
        EnsureMemoEntriesColumn(connection, "Tags", "TEXT");

        using SqliteCommand createImageTable = connection.CreateCommand();
        createImageTable.CommandText = """
            CREATE TABLE IF NOT EXISTS MemoImages (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                MemoEntryId INTEGER NOT NULL,
                ImageData BLOB NOT NULL,
                FOREIGN KEY (MemoEntryId) REFERENCES MemoEntries(Id) ON DELETE CASCADE
            );
            """;
        createImageTable.ExecuteNonQuery();

        using SqliteTransaction transaction = connection.BeginTransaction();
        foreach (MemoRow emptyExistingRow in _rows.Where(row => row.Id > 0 && !HasMeaningfulData(row)).ToList())
        {
            _deletedRowIds.Add(emptyExistingRow.Id);
        }

        foreach (long deletedRowId in _deletedRowIds)
        {
            using SqliteCommand deleteImages = connection.CreateCommand();
            deleteImages.Transaction = transaction;
            deleteImages.CommandText = "DELETE FROM MemoImages WHERE MemoEntryId = @memoEntryId;";
            deleteImages.Parameters.AddWithValue("@memoEntryId", deletedRowId);
            deleteImages.ExecuteNonQuery();

            using SqliteCommand deleteRow = connection.CreateCommand();
            deleteRow.Transaction = transaction;
            deleteRow.CommandText = "DELETE FROM MemoEntries WHERE Id = @id;";
            deleteRow.Parameters.AddWithValue("@id", deletedRowId);
            deleteRow.ExecuteNonQuery();
        }

        foreach (MemoRow row in _rows.Where(HasMeaningfulData))
        {
            long memoEntryId = row.Id;
            if (row.Id <= 0)
            {
                using SqliteCommand insertMemo = connection.CreateCommand();
                insertMemo.Transaction = transaction;
                insertMemo.CommandText = """
                    INSERT INTO MemoEntries (MemoDate, Model, Category, Tags, Issue, CreatedAtUtc)
                    VALUES (@memoDate, @model, @category, @tags, @issue, @createdAtUtc);
                    SELECT last_insert_rowid();
                    """;
                insertMemo.Parameters.AddWithValue("@memoDate", row.MemoDate.Trim());
                insertMemo.Parameters.AddWithValue("@model", row.Model.Trim());
                insertMemo.Parameters.AddWithValue("@category", row.Category.Trim());
                insertMemo.Parameters.AddWithValue("@tags", row.Tags.Trim());
                insertMemo.Parameters.AddWithValue("@issue", row.Issue.Trim());
                insertMemo.Parameters.AddWithValue("@createdAtUtc", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
                memoEntryId = (long)(insertMemo.ExecuteScalar() ?? 0L);
                row.Id = memoEntryId;
            }
            else if (row.IsDirty)
            {
                using SqliteCommand updateMemo = connection.CreateCommand();
                updateMemo.Transaction = transaction;
                updateMemo.CommandText = """
                    UPDATE MemoEntries
                    SET MemoDate = @memoDate,
                        Model = @model,
                        Category = @category,
                        Tags = @tags,
                        Issue = @issue
                    WHERE Id = @id;
                    """;
                updateMemo.Parameters.AddWithValue("@id", row.Id);
                updateMemo.Parameters.AddWithValue("@memoDate", row.MemoDate.Trim());
                updateMemo.Parameters.AddWithValue("@model", row.Model.Trim());
                updateMemo.Parameters.AddWithValue("@category", row.Category.Trim());
                updateMemo.Parameters.AddWithValue("@tags", row.Tags.Trim());
                updateMemo.Parameters.AddWithValue("@issue", row.Issue.Trim());
                updateMemo.ExecuteNonQuery();
            }

            foreach (long deletedImageId in row.DeletedImageIds)
            {
                using SqliteCommand deleteImage = connection.CreateCommand();
                deleteImage.Transaction = transaction;
                deleteImage.CommandText = "DELETE FROM MemoImages WHERE Id = @id;";
                deleteImage.Parameters.AddWithValue("@id", deletedImageId);
                deleteImage.ExecuteNonQuery();
            }

            foreach (MemoImageItem image in row.Images)
            {
                if (image.Id > 0)
                {
                    continue;
                }

                using SqliteCommand insertImage = connection.CreateCommand();
                insertImage.Transaction = transaction;
                insertImage.CommandText = """
                    INSERT INTO MemoImages (MemoEntryId, ImageData)
                    VALUES (@memoEntryId, @imageData);
                    SELECT last_insert_rowid();
                    """;
                insertImage.Parameters.AddWithValue("@memoEntryId", memoEntryId);
                insertImage.Parameters.Add("@imageData", SqliteType.Blob).Value = image.ImageBytes;
                image.Id = (long)(insertImage.ExecuteScalar() ?? 0L);
            }

            row.MarkClean();
        }

        transaction.Commit();
        _deletedRowIds.Clear();
    }

    private void LoadRowsFromCurrentDatabase()
    {
        string databasePath = LoadSavedDatabasePath();
        _rows.Clear();
        _deletedRowIds.Clear();

        if (string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
        {
            AddNewRow();
            StatusTextBlock.Text = "Memo DB loaded. No saved rows found.";
            NotifyWebModuleSnapshotChanged();
            return;
        }

        try
        {
            LoadRowsFromDatabase(databasePath);
            if (_rows.Count == 0)
            {
                AddNewRow();
            }

            StatusTextBlock.Text = $"Loaded {_rows.Count:N0} memo row(s) from DB.";
            _selectedRow ??= _rows.FirstOrDefault();
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            _rows.Clear();
            AddNewRow();
            StatusTextBlock.Text = $"Failed to load memo DB: {ex.Message}";
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void LoadRowsFromDatabase(string databasePath)
    {
        using SqliteConnection connection = new($"Data Source={databasePath}");
        connection.Open();

        if (!TableExists(connection, "MemoEntries"))
        {
            return;
        }

        bool hasTagsColumn = TableColumnExists(connection, "MemoEntries", "Tags");

        using SqliteCommand loadMemoCommand = connection.CreateCommand();
        loadMemoCommand.CommandText = hasTagsColumn
            ? """
              SELECT Id, MemoDate, Model, Category, Tags, Issue
              FROM MemoEntries
              ORDER BY Id;
              """
            : """
              SELECT Id, MemoDate, Model, Category, Issue
              FROM MemoEntries
              ORDER BY Id;
              """;

        using SqliteDataReader reader = loadMemoCommand.ExecuteReader();
        while (reader.Read())
        {
            long memoEntryId = reader.GetInt64(0);
            MemoRow row = new()
            {
                Id = memoEntryId,
                MemoDate = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                Model = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                Category = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                Tags = hasTagsColumn || reader.FieldCount > 5
                    ? (reader.IsDBNull(4) ? string.Empty : reader.GetString(4))
                    : string.Empty,
                Issue = hasTagsColumn || reader.FieldCount > 5
                    ? (reader.IsDBNull(5) ? string.Empty : reader.GetString(5))
                    : (reader.IsDBNull(4) ? string.Empty : reader.GetString(4))
            };

            foreach ((long imageId, byte[] imageBytes) in LoadImagesForMemoEntry(connection, memoEntryId))
            {
                row.AddLoadedImage(imageId, imageBytes);
            }

            row.MarkClean();
            _rows.Add(row);
        }
    }

    public object GetWebModuleSnapshot()
    {
        string databasePath = LoadSavedDatabasePath();
        MemoRow? selectedRow = _selectedRow ?? _rows.FirstOrDefault();

        return new
        {
            moduleType = "Memo",
            databasePath,
            statusMessage = StatusTextBlock.Text ?? string.Empty,
            rowCount = _rows.Count,
            selectedRowId = selectedRow?.Id ?? 0,
            rows = _rows.Take(120).Select(row => new
            {
                id = row.Id,
                memoDate = row.MemoDate,
                model = row.Model,
                category = row.Category,
                tags = row.TagItems.ToArray(),
                issue = row.Issue,
                imageCount = row.Images.Count,
                isSelected = ReferenceEquals(row, selectedRow)
            }).ToArray()
        };
    }

    public object UpdateWebModuleState(JsonElement payload)
    {
        if (payload.ValueKind != JsonValueKind.Object)
        {
            return GetWebModuleSnapshot();
        }

        if (payload.TryGetProperty("selectedRowId", out JsonElement selectedRowIdElement)
            && selectedRowIdElement.TryGetInt64(out long selectedRowId))
        {
            MemoRow? row = _rows.FirstOrDefault(item => item.Id == selectedRowId)
                ?? _rows.ElementAtOrDefault((int)Math.Max(0, selectedRowId));
            if (row is not null)
            {
                _selectedRow = row;
                MemoDataGrid.SelectedItem = row;
                MemoDataGrid.CurrentItem = row;
                MemoDataGrid.ScrollIntoView(row);
                StatusTextBlock.Text = "Memo row selected.";
            }
        }

        return GetWebModuleSnapshot();
    }

    public object InvokeWebModuleAction(string action)
    {
        switch (action)
        {
            case "add-row":
                AddRowButton_Click(this, new RoutedEventArgs());
                break;
            case "delete-row":
                DeleteRowButton_Click(this, new RoutedEventArgs());
                break;
            case "save-memo":
                SaveButton_Click(this, new RoutedEventArgs());
                break;
        }

        return GetWebModuleSnapshot();
    }

    private void NotifyWebModuleSnapshotChanged()
    {
        WebModuleSnapshotChanged?.Invoke();
    }

    private static bool TableExists(SqliteConnection connection, string tableName)
    {
        using SqliteCommand command = connection.CreateCommand();
        command.CommandText = """
            SELECT COUNT(1)
            FROM sqlite_master
            WHERE type = 'table' AND name = @tableName;
            """;
        command.Parameters.AddWithValue("@tableName", tableName);
        return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture) > 0;
    }

    private static void EnsureMemoEntriesColumn(SqliteConnection connection, string columnName, string columnType)
    {
        if (TableColumnExists(connection, "MemoEntries", columnName))
        {
            return;
        }

        using SqliteCommand alterCommand = connection.CreateCommand();
        alterCommand.CommandText = $"ALTER TABLE MemoEntries ADD COLUMN {columnName} {columnType};";
        alterCommand.ExecuteNonQuery();
    }

    private static bool TableColumnExists(SqliteConnection connection, string tableName, string columnName)
    {
        using SqliteCommand command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({tableName});";

        using SqliteDataReader reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader["name"]?.ToString(), columnName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static List<(long Id, byte[] ImageBytes)> LoadImagesForMemoEntry(SqliteConnection connection, long memoEntryId)
    {
        if (!TableExists(connection, "MemoImages"))
        {
            return [];
        }

        using SqliteCommand command = connection.CreateCommand();
        command.CommandText = """
            SELECT Id, ImageData
            FROM MemoImages
            WHERE MemoEntryId = @memoEntryId
            ORDER BY Id;
            """;
        command.Parameters.AddWithValue("@memoEntryId", memoEntryId);

        using SqliteDataReader reader = command.ExecuteReader();
        List<(long Id, byte[] ImageBytes)> images = [];
        while (reader.Read())
        {
            if (!reader.IsDBNull(0) && !reader.IsDBNull(1))
            {
                images.Add((reader.GetInt64(0), (byte[])reader["ImageData"]));
            }
        }

        return images;
    }

    private static bool HasMeaningfulData(MemoRow row)
    {
        return !string.IsNullOrWhiteSpace(row.MemoDate) ||
               !string.IsNullOrWhiteSpace(row.Model) ||
               !string.IsNullOrWhiteSpace(row.Category) ||
               !string.IsNullOrWhiteSpace(row.Tags) ||
               !string.IsNullOrWhiteSpace(row.Issue) ||
               row.Images.Count > 0;
    }

    private MemoRow? FindOwningMemoRow(MemoImageItem imageItem)
    {
        return _rows.FirstOrDefault(row => row.Images.Contains(imageItem));
    }

    private void SortRowsByTag(string tag)
    {
        List<MemoRow> orderedRows = _rows
            .OrderByDescending(row => row.TagItems.Contains(tag, StringComparer.OrdinalIgnoreCase))
            .ThenBy(row => row.MemoDate, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.Model, StringComparer.OrdinalIgnoreCase)
            .ToList();

        _rows.Clear();
        foreach (MemoRow row in orderedRows)
        {
            _rows.Add(row);
        }
    }
}
