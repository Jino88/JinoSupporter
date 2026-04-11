using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 파일 드래그앤드롭 + Browse 버튼을 제공하는 공통 파일 입력 컨트롤.
    /// 파일이 선택되면 <see cref="FilesSelected"/> 이벤트로 경로 배열을 전달한다.
    /// </summary>
    public partial class AppFileDropBox : System.Windows.Controls.UserControl
    {
        // ── 기본 색상 ────────────────────────────────────────────────────────
        private static readonly SolidColorBrush _defaultBorder  = new(Color.FromRgb(200, 214, 232));
        private static readonly SolidColorBrush _hoverBorder    = new(Color.FromRgb(10,  99,  201));
        private static readonly SolidColorBrush _defaultHint    = new(Color.FromRgb(100, 116, 139));
        private static readonly SolidColorBrush _hoverHint      = new(Color.FromRgb(10,  99,  201));
        private static readonly SolidColorBrush _defaultBg      = new(Color.FromRgb(250, 251, 254));
        private static readonly SolidColorBrush _hoverBg        = new(Color.FromRgb(235, 243, 255));

        // ─────────────────────────────────────────────────────────────────────
        //  DependencyProperties
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>드롭 존에 표시할 힌트 텍스트. 기본값: "Drag data files"</summary>
        public static readonly DependencyProperty HintTextProperty =
            DependencyProperty.Register(nameof(HintText), typeof(string), typeof(AppFileDropBox),
                new PropertyMetadata("Drag data files", OnHintTextChanged));

        public string HintText
        {
            get => (string)GetValue(HintTextProperty);
            set => SetValue(HintTextProperty, value);
        }

        /// <summary>Browse 버튼에 표시할 텍스트. 기본값: "Browse"</summary>
        public static readonly DependencyProperty BrowseButtonTextProperty =
            DependencyProperty.Register(nameof(BrowseButtonText), typeof(string), typeof(AppFileDropBox),
                new PropertyMetadata("Browse", OnBrowseButtonTextChanged));

        public string BrowseButtonText
        {
            get => (string)GetValue(BrowseButtonTextProperty);
            set => SetValue(BrowseButtonTextProperty, value);
        }

        /// <summary>허용할 파일 확장자 배열. 기본값: [".txt", ".csv"]</summary>
        public static readonly DependencyProperty AllowedExtensionsProperty =
            DependencyProperty.Register(nameof(AllowedExtensions), typeof(string[]), typeof(AppFileDropBox),
                new PropertyMetadata(new[] { ".txt", ".csv" }));

        public string[] AllowedExtensions
        {
            get => (string[])GetValue(AllowedExtensionsProperty);
            set => SetValue(AllowedExtensionsProperty, value);
        }

        /// <summary>여러 파일 선택 허용 여부. 기본값: true</summary>
        public static readonly DependencyProperty AllowMultipleProperty =
            DependencyProperty.Register(nameof(AllowMultiple), typeof(bool), typeof(AppFileDropBox),
                new PropertyMetadata(true));

        public bool AllowMultiple
        {
            get => (bool)GetValue(AllowMultipleProperty);
            set => SetValue(AllowMultipleProperty, value);
        }

        /// <summary>
        /// OpenFileDialog의 Filter 문자열.
        /// 기본값: "Text/CSV files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*"
        /// </summary>
        public static readonly DependencyProperty FileFilterProperty =
            DependencyProperty.Register(nameof(FileFilter), typeof(string), typeof(AppFileDropBox),
                new PropertyMetadata("Text/CSV files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*"));

        public string FileFilter
        {
            get => (string)GetValue(FileFilterProperty);
            set => SetValue(FileFilterProperty, value);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  이벤트
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// 파일이 선택됐을 때 발생. 드래그앤드롭과 Browse 버튼 모두 이 이벤트를 통해 전달됨.
        /// EventArgs에 선택된 파일 경로 배열이 담겨 있음.
        /// </summary>
        public event EventHandler<FilesSelectedEventArgs>? FilesSelected;

        // ─────────────────────────────────────────────────────────────────────
        //  생성자
        // ─────────────────────────────────────────────────────────────────────

        public AppFileDropBox()
        {
            InitializeComponent();
            HintTextBlock.Text    = HintText;
            BrowseButton.Content  = BrowseButtonText;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  PropertyChanged 콜백
        // ─────────────────────────────────────────────────────────────────────

        private static void OnHintTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is AppFileDropBox ctrl)
                ctrl.HintTextBlock.Text = (string)e.NewValue;
        }

        private static void OnBrowseButtonTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is AppFileDropBox ctrl)
                ctrl.BrowseButton.Content = (string)e.NewValue;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Browse 버튼
        // ─────────────────────────────────────────────────────────────────────

        private void BrowseButton_Click(object sender, RoutedEventArgs e) => OnBrowseClicked();

        /// <summary>코드에서 Browse 다이얼로그를 직접 열 때 사용.</summary>
        public void Browse() => OnBrowseClicked();

        /// <summary>Browse 버튼 클릭 시 동작. 파생 클래스에서 override 가능.</summary>
        protected virtual void OnBrowseClicked()
        {
            var dialog = new OpenFileDialog
            {
                Title       = "Select file",
                Filter      = FileFilter,
                Multiselect = AllowMultiple
            };

            if (dialog.ShowDialog() == true)
                RaiseFilesSelected(dialog.FileNames);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  드래그앤드롭 이벤트
        // ─────────────────────────────────────────────────────────────────────

        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.None;
                return;
            }
            e.Effects = DragDropEffects.Copy;
            SetDropZoneState(isHover: true);
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            SetDropZoneState(isHover: false);
        }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            SetDropZoneState(isHover: false);

            if (e.Data.GetData(DataFormats.FileDrop) is string[] droppedPaths)
                OnDropped(droppedPaths);
        }

        /// <summary>드롭 완료 시 동작. 파생 클래스에서 override 가능.</summary>
        protected virtual void OnDropped(string[] droppedPaths)
        {
            var valid = droppedPaths
                .Where(p => IsAllowedExtension(p))
                .ToArray();

            if (valid.Length == 0)
                return;

            RaiseFilesSelected(AllowMultiple ? valid : new[] { valid[0] });
        }

        // ─────────────────────────────────────────────────────────────────────
        //  파일 내용 읽기 — 정적 헬퍼
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// 파일 경로를 받아 전체 줄 배열을 반환한다.
        /// 인코딩 자동 감지 (BOM 우선, 없으면 UTF-8).
        /// </summary>
        public static string[] ReadAllLines(string filePath)
        {
            using var reader = new StreamReader(filePath, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd()
                         .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        }

        /// <summary>
        /// 파일 경로를 받아 전체 텍스트를 반환한다.
        /// </summary>
        public static string ReadAllText(string filePath)
        {
            using var reader = new StreamReader(filePath, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }

        // ─────────────────────────────────────────────────────────────────────
        //  내부 헬퍼
        // ─────────────────────────────────────────────────────────────────────

        protected void RaiseFilesSelected(string[] paths)
        {
            FilesSelected?.Invoke(this, new FilesSelectedEventArgs(paths));
        }

        private bool IsAllowedExtension(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return AllowedExtensions.Length == 0
                || AllowedExtensions.Any(a => a.ToLowerInvariant() == ext);
        }

        private void SetDropZoneState(bool isHover)
        {
            DropZoneBorder.BorderBrush  = isHover ? _hoverBorder  : _defaultBorder;
            DropZoneBorder.Background   = isHover ? _hoverBg      : _defaultBg;
            HintTextBlock.Foreground    = isHover ? _hoverHint    : _defaultHint;
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  EventArgs
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>파일 선택 이벤트 데이터.</summary>
    public sealed class FilesSelectedEventArgs : EventArgs
    {
        /// <summary>선택된 파일 경로 배열.</summary>
        public string[] FilePaths { get; }

        public FilesSelectedEventArgs(string[] filePaths)
        {
            FilePaths = filePaths;
        }
    }
}
