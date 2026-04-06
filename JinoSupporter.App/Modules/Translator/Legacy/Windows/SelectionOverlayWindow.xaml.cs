using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace CustomKeyboardCSharp.Windows;

public partial class SelectionOverlayWindow : Window
{
    private readonly System.Drawing.Rectangle _screenBounds;
    private System.Windows.Point _startPoint;
    private bool _isSelecting;

    public SelectionOverlayWindow(BitmapSource screenshot, System.Drawing.Rectangle screenBounds)
    {
        InitializeComponent();
        _screenBounds = screenBounds;
        Left = screenBounds.Left;
        Top = screenBounds.Top;
        Width = screenBounds.Width;
        Height = screenBounds.Height;
        ScreenshotImage.Source = screenshot;
        PreviewKeyDown += OnPreviewKeyDown;
    }

    public Rect? SelectedRegion { get; private set; }

    private void OverlayCanvas_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _startPoint = e.GetPosition(OverlayCanvas);
        _isSelecting = true;
        SelectionRectangle.Visibility = Visibility.Visible;
        Canvas.SetLeft(SelectionRectangle, _startPoint.X);
        Canvas.SetTop(SelectionRectangle, _startPoint.Y);
        SelectionRectangle.Width = 0;
        SelectionRectangle.Height = 0;
        OverlayCanvas.CaptureMouse();
    }

    private void OverlayCanvas_OnMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isSelecting)
        {
            return;
        }

        var current = e.GetPosition(OverlayCanvas);
        var x = Math.Min(current.X, _startPoint.X);
        var y = Math.Min(current.Y, _startPoint.Y);
        var width = Math.Abs(current.X - _startPoint.X);
        var height = Math.Abs(current.Y - _startPoint.Y);

        Canvas.SetLeft(SelectionRectangle, x);
        Canvas.SetTop(SelectionRectangle, y);
        SelectionRectangle.Width = width;
        SelectionRectangle.Height = height;
    }

    private void OverlayCanvas_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isSelecting)
        {
            return;
        }

        _isSelecting = false;
        OverlayCanvas.ReleaseMouseCapture();

        var x = Canvas.GetLeft(SelectionRectangle);
        var y = Canvas.GetTop(SelectionRectangle);
        var width = SelectionRectangle.Width;
        var height = SelectionRectangle.Height;

        if (width < 4 || height < 4)
        {
            DialogResult = false;
            return;
        }

        SelectedRegion = new Rect(
            x + _screenBounds.Left,
            y + _screenBounds.Top,
            width,
            height);

        DialogResult = true;
    }

    private void OnPreviewKeyDown(object? sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            DialogResult = false;
        }
    }
}
