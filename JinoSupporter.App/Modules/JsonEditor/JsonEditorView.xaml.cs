using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WorkbenchHost.Modules.JsonEditor;

public partial class JsonEditorView : UserControl, INotifyPropertyChanged
{
    private JsonTreeNode? _rootNode;
    private JsonTreeNode? _selectedNode;
    private string _currentFilePath = "No file selected";
    private string _statusMessage = "JSON file is not loaded.";
    private bool _isDirty;

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action? WebModuleSnapshotChanged;

    public string CurrentFilePath
    {
        get => _currentFilePath;
        private set
        {
            if (_currentFilePath == value)
            {
                return;
            }

            _currentFilePath = value;
            OnPropertyChanged();
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set
        {
            if (_statusMessage == value)
            {
                return;
            }

            _statusMessage = value;
            OnPropertyChanged();
        }
    }

    public string DirtyStateText => _isDirty ? "Modified" : "Saved";

    public JsonEditorView()
    {
        InitializeComponent();
        DataContext = this;
    }

    private void OpenJsonButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            LoadJsonFile(dialog.FileName);
        }
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        if (_rootNode is null || !File.Exists(CurrentFilePath))
        {
            return;
        }

        SaveJsonToPath(CurrentFilePath);
    }

    private void SaveAsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_rootNode is null)
        {
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            DefaultExt = ".json",
            FileName = Path.GetFileName(CurrentFilePath) is { Length: > 0 } name && name != "No file selected" ? name : "data.json"
        };

        if (dialog.ShowDialog() == true)
        {
            SaveJsonToPath(dialog.FileName);
        }
    }

    private void ReloadButton_Click(object sender, RoutedEventArgs e)
    {
        if (!File.Exists(CurrentFilePath))
        {
            return;
        }

        LoadJsonFile(CurrentFilePath);
    }

    private void UserControl_DragOver(object sender, DragEventArgs e)
    {
        if (TryGetDraggedJsonPath(e.Data) is not null)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
            return;
        }

        e.Effects = DragDropEffects.None;
        e.Handled = true;
    }

    private void UserControl_Drop(object sender, DragEventArgs e)
    {
        string? path = TryGetDraggedJsonPath(e.Data);
        if (path is null)
        {
            return;
        }

        LoadJsonFile(path);
    }

    private void JsonTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is JsonTreeNode node)
        {
            _selectedNode = node;
            StatusMessage = $"Selected: {node.DisplayPath}";
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void AddPropertyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: JsonTreeNode parent })
        {
            return;
        }

        var dialog = new JsonNodeDialog(requiresKey: true) { Owner = Window.GetWindow(this) };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        if (parent.Children.Any(child => string.Equals(child.Key, dialog.NodeKey, StringComparison.Ordinal)))
        {
            MessageBox.Show(Window.GetWindow(this), "A property with the same key already exists.", "JsonEditor", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        JsonTreeNode child = JsonTreeNode.CreateNew(dialog.NodeKey, dialog.NodeKind, dialog.NodeValue, parent);
        parent.Children.Add(child);
        parent.IsExpanded = true;
        AttachNodeRecursive(child);
        MarkDirty($"Added property '{dialog.NodeKey}'.");
        NotifyWebModuleSnapshotChanged();
    }

    private void AddItemButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: JsonTreeNode parent })
        {
            return;
        }

        var dialog = new JsonNodeDialog(requiresKey: false) { Owner = Window.GetWindow(this) };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        JsonTreeNode child = JsonTreeNode.CreateNew(parent.Children.Count.ToString(), dialog.NodeKind, dialog.NodeValue, parent);
        parent.Children.Add(child);
        parent.IsExpanded = true;
        AttachNodeRecursive(child);
        ReindexArrayChildren(parent);
        MarkDirty("Added array item.");
        NotifyWebModuleSnapshotChanged();
    }

    private void DeleteNodeButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: JsonTreeNode node } || node.Parent is null)
        {
            return;
        }

        MessageBoxResult result = MessageBox.Show(Window.GetWindow(this),
            $"Delete '{node.DisplayName}'?",
            "JsonEditor",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        JsonTreeNode parent = node.Parent;
        parent.Children.Remove(node);
        if (parent.Kind == JsonTreeNodeKind.Array)
        {
            ReindexArrayChildren(parent);
        }

        MarkDirty($"Deleted '{node.DisplayName}'.");
        NotifyWebModuleSnapshotChanged();
    }

    private void LoadJsonFile(string path)
    {
        if (!ConfirmDiscardChanges())
        {
            return;
        }

        try
        {
            string json = File.ReadAllText(path, Encoding.UTF8);
            JToken token = JToken.Parse(json);
            _rootNode = JsonTreeNode.FromJToken("Root", token, null, true);
            AttachNodeRecursive(_rootNode);
            JsonTreeView.ItemsSource = new[] { _rootNode };
            _selectedNode = _rootNode;
            CurrentFilePath = path;
            _isDirty = false;
            OnPropertyChanged(nameof(DirtyStateText));
            StatusMessage = $"Loaded {Path.GetFileName(path)}";
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            MessageBox.Show(Window.GetWindow(this), ex.Message, "JsonEditor", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = $"Load failed: {Path.GetFileName(path)}";
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void SaveJsonToPath(string path)
    {
        try
        {
            if (_rootNode is null)
            {
                return;
            }

            JToken token = _rootNode.ToJToken();
            File.WriteAllText(path, token.ToString(Formatting.Indented), new UTF8Encoding(false));
            CurrentFilePath = path;
            _isDirty = false;
            OnPropertyChanged(nameof(DirtyStateText));
            StatusMessage = $"Saved {Path.GetFileName(path)}";
            NotifyWebModuleSnapshotChanged();
        }
        catch (Exception ex)
        {
            MessageBox.Show(Window.GetWindow(this), ex.Message, "JsonEditor", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusMessage = "Save failed.";
            NotifyWebModuleSnapshotChanged();
        }
    }

    private bool ConfirmDiscardChanges()
    {
        if (!_isDirty)
        {
            return true;
        }

        return MessageBox.Show(Window.GetWindow(this),
            "Unsaved changes will be lost. Continue?",
            "JsonEditor",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning) == MessageBoxResult.Yes;
    }

    private void AttachNodeRecursive(JsonTreeNode node)
    {
        node.Modified -= Node_Modified;
        node.Modified += Node_Modified;
        node.Children.CollectionChanged -= NodeChildren_CollectionChanged;
        node.Children.CollectionChanged += NodeChildren_CollectionChanged;

        foreach (JsonTreeNode child in node.Children)
        {
            AttachNodeRecursive(child);
        }
    }

    private void NodeChildren_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (JsonTreeNode child in e.NewItems.OfType<JsonTreeNode>())
            {
                AttachNodeRecursive(child);
            }
        }
    }

    private void Node_Modified(object? sender, EventArgs e)
    {
        MarkDirty("JSON modified.");
    }

    private void MarkDirty(string message)
    {
        _isDirty = true;
        OnPropertyChanged(nameof(DirtyStateText));
        StatusMessage = message;
        NotifyWebModuleSnapshotChanged();
    }

    private static string? TryGetDraggedJsonPath(IDataObject dataObject)
    {
        if (!dataObject.GetDataPresent(DataFormats.FileDrop))
        {
            return null;
        }

        if (dataObject.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0)
        {
            return null;
        }

        string path = files[0];
        return string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase) && File.Exists(path)
            ? path
            : null;
    }

    private static void ReindexArrayChildren(JsonTreeNode arrayNode)
    {
        for (int i = 0; i < arrayNode.Children.Count; i++)
        {
            arrayNode.Children[i].SetKeyWithoutNotification(i.ToString());
        }
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public object GetWebModuleSnapshot()
    {
        return new
        {
            moduleType = "JsonEditor",
            currentFilePath = CurrentFilePath,
            statusMessage = StatusMessage,
            dirtyState = DirtyStateText,
            hasDocument = _rootNode is not null,
            selectedPath = _selectedNode?.DisplayPath ?? string.Empty,
            nodes = _rootNode is null
                ? Array.Empty<object>()
                : new[] { BuildNodeSnapshot(_rootNode, 0) }
        };
    }

    public object UpdateWebModuleState(JsonElement payload)
    {
        if (payload.ValueKind != JsonValueKind.Object)
        {
            return GetWebModuleSnapshot();
        }

        if (payload.TryGetProperty("selectedPath", out JsonElement selectedPathElement))
        {
            string selectedPath = selectedPathElement.GetString() ?? string.Empty;
            JsonTreeNode? node = FindNodeByPath(_rootNode, selectedPath);
            if (node is not null)
            {
                _selectedNode = node;
                StatusMessage = $"Selected: {node.DisplayPath}";
            }
        }

        return GetWebModuleSnapshot();
    }

    public Task<object> InvokeWebModuleActionAsync(string action)
    {
        switch (action)
        {
            case "open-json":
                OpenJsonButton_Click(this, new RoutedEventArgs());
                break;
            case "save-json":
                SaveButton_Click(this, new RoutedEventArgs());
                break;
            case "save-json-as":
                SaveAsButton_Click(this, new RoutedEventArgs());
                break;
            case "reload-json":
                ReloadButton_Click(this, new RoutedEventArgs());
                break;
        }

        return Task.FromResult(GetWebModuleSnapshot());
    }

    private void NotifyWebModuleSnapshotChanged()
    {
        WebModuleSnapshotChanged?.Invoke();
    }

    private static object BuildNodeSnapshot(JsonTreeNode node, int depth)
    {
        return new
        {
            key = node.DisplayName,
            rawKey = node.Key,
            kind = node.KindLabel,
            value = node.IsContainer ? node.EditableValue : node.EditableValue,
            path = node.DisplayPath,
            depth,
            canAddProperty = node.CanAddProperty,
            canAddItem = node.CanAddItem,
            canDelete = node.CanDelete,
            children = node.Children.Select(child => BuildNodeSnapshot(child, depth + 1)).ToArray()
        };
    }

    private static JsonTreeNode? FindNodeByPath(JsonTreeNode? rootNode, string path)
    {
        if (rootNode is null || string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        if (string.Equals(rootNode.DisplayPath, path, StringComparison.Ordinal))
        {
            return rootNode;
        }

        foreach (JsonTreeNode child in rootNode.Children)
        {
            JsonTreeNode? match = FindNodeByPath(child, path);
            if (match is not null)
            {
                return match;
            }
        }

        return null;
    }
}

public enum JsonTreeNodeKind
{
    Object,
    Array,
    String,
    Number,
    Boolean,
    Null
}

public sealed class JsonTreeNode : INotifyPropertyChanged
{
    private static readonly string[] LockedNodeNames = ["Index", "ModelList"];
    private static readonly System.Windows.Media.Brush RootBrush =
        new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F3F6FB"));
    private static readonly System.Windows.Media.Brush ContainerBrush =
        new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F7F9FC"));

    private string _key = string.Empty;
    private string _editableValue = string.Empty;
    private bool _isExpanded;

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler? Modified;

    public ObservableCollection<JsonTreeNode> Children { get; } = new();

    public JsonTreeNode? Parent { get; set; }
    public bool IsRoot { get; set; }
    public JsonTreeNodeKind Kind { get; set; }

    public string Key
    {
        get => _key;
        set
        {
            if (_key == value)
            {
                return;
            }

            _key = value;
            NotifyIdentityChanged();
            RaiseModified();
        }
    }

    public string EditableKey
    {
        get => Key;
        set => Key = value;
    }

    public string EditableValue
    {
        get => _editableValue;
        set
        {
            if (_editableValue == value)
            {
                return;
            }

            _editableValue = value;
            OnPropertyChanged();
            RaiseModified();
        }
    }

    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            if (_isExpanded == value)
            {
                return;
            }

            _isExpanded = value;
            OnPropertyChanged();
        }
    }

    public bool IsContainer => Kind is JsonTreeNodeKind.Object or JsonTreeNodeKind.Array;
    public bool IsLockedNode => LockedNodeNames.Any(name => string.Equals(name, Key, StringComparison.OrdinalIgnoreCase));
    public bool IsKeyReadOnly => IsRoot || IsLockedNode;
    public bool IsValueReadOnly => IsContainer || IsLockedNode;
    public bool CanAddProperty => Kind == JsonTreeNodeKind.Object;
    public bool CanAddItem => Kind == JsonTreeNodeKind.Array;
    public bool CanDelete => !IsRoot;
    public string KindLabel => Kind.ToString();
    public string DisplayName => IsRoot ? "Root" : (Parent?.Kind == JsonTreeNodeKind.Array ? $"[{Key}]" : Key);
    public string DisplayPath => Parent is null ? DisplayName : $"{Parent.DisplayPath}/{DisplayName}";
    public System.Windows.Media.Brush KeyBackground => IsRoot ? RootBrush : System.Windows.Media.Brushes.White;
    public System.Windows.Media.Brush ValueBackground => IsContainer ? ContainerBrush : System.Windows.Media.Brushes.White;

    public static JsonTreeNode FromJToken(string key, JToken token, JsonTreeNode? parent, bool isRoot = false)
    {
        var node = new JsonTreeNode
        {
            Parent = parent,
            IsRoot = isRoot,
            _key = key
        };

        switch (token.Type)
        {
            case JTokenType.Object:
                node.Kind = JsonTreeNodeKind.Object;
                node._editableValue = "{ }";
                node._isExpanded = true;
                foreach (JProperty property in token.Children<JProperty>())
                {
                    node.Children.Add(FromJToken(property.Name, property.Value, node));
                }
                break;
            case JTokenType.Array:
                node.Kind = JsonTreeNodeKind.Array;
                node._editableValue = "[ ]";
                node._isExpanded = true;
                int index = 0;
                foreach (JToken child in token.Children())
                {
                    node.Children.Add(FromJToken(index.ToString(), child, node));
                    index++;
                }
                break;
            case JTokenType.Integer:
            case JTokenType.Float:
                node.Kind = JsonTreeNodeKind.Number;
                node._editableValue = token.ToString(Formatting.None);
                break;
            case JTokenType.Boolean:
                node.Kind = JsonTreeNodeKind.Boolean;
                node._editableValue = token.Value<bool>() ? "true" : "false";
                break;
            case JTokenType.Null:
            case JTokenType.Undefined:
                node.Kind = JsonTreeNodeKind.Null;
                node._editableValue = string.Empty;
                break;
            default:
                node.Kind = JsonTreeNodeKind.String;
                node._editableValue = token.ToString();
                break;
        }

        return node;
    }

    public static JsonTreeNode CreateNew(string key, JsonTreeNodeKind kind, string rawValue, JsonTreeNode parent)
    {
        return new JsonTreeNode
        {
            Parent = parent,
            _key = key,
            Kind = kind,
            _isExpanded = kind is JsonTreeNodeKind.Object or JsonTreeNodeKind.Array,
            _editableValue = kind switch
            {
                JsonTreeNodeKind.Object => "{ }",
                JsonTreeNodeKind.Array => "[ ]",
                _ => rawValue
            }
        };
    }

    public JToken ToJToken()
    {
        return Kind switch
        {
            JsonTreeNodeKind.Object => new JObject(Children.Select(child => new JProperty(child.Key, child.ToJToken()))),
            JsonTreeNodeKind.Array => new JArray(Children.Select(child => child.ToJToken())),
            JsonTreeNodeKind.String => new JValue(EditableValue),
            JsonTreeNodeKind.Number => new JValue(ParseNumber()),
            JsonTreeNodeKind.Boolean => new JValue(ParseBoolean()),
            JsonTreeNodeKind.Null => JValue.CreateNull(),
            _ => new JValue(EditableValue)
        };
    }

    public void SetKeyWithoutNotification(string key)
    {
        _key = key;
        NotifyIdentityChanged();
    }

    private decimal ParseNumber()
    {
        if (decimal.TryParse(EditableValue, out decimal value))
        {
            return value;
        }

        throw new InvalidOperationException($"'{EditableValue}' is not a valid number.");
    }

    private bool ParseBoolean()
    {
        if (bool.TryParse(EditableValue, out bool value))
        {
            return value;
        }

        throw new InvalidOperationException($"'{EditableValue}' is not a valid boolean.");
    }

    private void NotifyIdentityChanged()
    {
        OnPropertyChanged(nameof(Key));
        OnPropertyChanged(nameof(EditableKey));
        OnPropertyChanged(nameof(DisplayName));
        OnPropertyChanged(nameof(DisplayPath));
    }

    private void RaiseModified()
    {
        Modified?.Invoke(this, EventArgs.Empty);
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
