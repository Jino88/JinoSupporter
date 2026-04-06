using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WorkbenchHost.Modules.Memo;

public partial class MemoImageViewerWindow : Window
{
    private const double FitScale = 1.0;
    private const double MinZoom = 0.2;
    private const double MaxZoom = 8.0;
    private const double ZoomStep = 1.15;

    private System.Windows.Point _dragStartPoint;
    private System.Windows.Point _dragStartTranslation;
    private bool _isDragging;

    public MemoImageViewerWindow(BitmapImage image)
    {
        InitializeComponent();
        PreviewImage.Source = image;
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ImageViewport_MouseWheel(object sender, MouseWheelEventArgs e)
    {
        double currentScale = ImageScaleTransform.ScaleX;
        double scaleFactor = e.Delta > 0 ? ZoomStep : 1 / ZoomStep;
        double targetScale = Math.Clamp(currentScale * scaleFactor, MinZoom, MaxZoom);
        if (Math.Abs(targetScale - currentScale) < double.Epsilon)
        {
            return;
        }

        System.Windows.Point cursorPosition = e.GetPosition(ImageViewport);
        double normalizedX = ImageViewport.ActualWidth <= 0 ? 0.5 : cursorPosition.X / ImageViewport.ActualWidth;
        double normalizedY = ImageViewport.ActualHeight <= 0 ? 0.5 : cursorPosition.Y / ImageViewport.ActualHeight;

        ImageScaleTransform.CenterX = (normalizedX - 0.5) * PreviewImage.ActualWidth;
        ImageScaleTransform.CenterY = (normalizedY - 0.5) * PreviewImage.ActualHeight;
        ImageScaleTransform.ScaleX = targetScale;
        ImageScaleTransform.ScaleY = targetScale;
        e.Handled = true;
    }

    private void ImageViewport_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            ToggleActualSize();
            e.Handled = true;
            return;
        }

        _isDragging = true;
        _dragStartPoint = e.GetPosition(ImageViewport);
        _dragStartTranslation = new System.Windows.Point(ImageTranslateTransform.X, ImageTranslateTransform.Y);
        ImageViewport.CaptureMouse();
        Cursor = Cursors.SizeAll;
    }

    private void ImageViewport_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        EndDrag();
    }

    private void ImageViewport_MouseMove(object sender, MouseEventArgs e)
    {
        if (!_isDragging)
        {
            return;
        }

        System.Windows.Point currentPoint = e.GetPosition(ImageViewport);
        Vector offset = currentPoint - _dragStartPoint;
        ImageTranslateTransform.X = _dragStartTranslation.X + offset.X;
        ImageTranslateTransform.Y = _dragStartTranslation.Y + offset.Y;
    }

    protected override void OnMouseLeave(MouseEventArgs e)
    {
        EndDrag();
        base.OnMouseLeave(e);
    }

    private void EndDrag()
    {
        if (!_isDragging)
        {
            return;
        }

        _isDragging = false;
        ImageViewport.ReleaseMouseCapture();
        Cursor = Cursors.Arrow;
    }

    private void ToggleActualSize()
    {
        bool isAtFitScale = Math.Abs(ImageScaleTransform.ScaleX - FitScale) < 0.001 &&
                            Math.Abs(ImageScaleTransform.ScaleY - FitScale) < 0.001;

        if (isAtFitScale)
        {
            double actualScale = GetActualSizeScale();
            ImageScaleTransform.CenterX = 0;
            ImageScaleTransform.CenterY = 0;
            ImageScaleTransform.ScaleX = actualScale;
            ImageScaleTransform.ScaleY = actualScale;
            ImageTranslateTransform.X = 0;
            ImageTranslateTransform.Y = 0;
            return;
        }

        ResetView();
    }

    private void ResetView()
    {
        ImageScaleTransform.CenterX = 0;
        ImageScaleTransform.CenterY = 0;
        ImageScaleTransform.ScaleX = FitScale;
        ImageScaleTransform.ScaleY = FitScale;
        ImageTranslateTransform.X = 0;
        ImageTranslateTransform.Y = 0;
    }

    private double GetActualSizeScale()
    {
        if (PreviewImage.Source is not BitmapSource bitmap ||
            ImageViewport.ActualWidth <= 0 ||
            ImageViewport.ActualHeight <= 0 ||
            bitmap.PixelWidth <= 0 ||
            bitmap.PixelHeight <= 0)
        {
            return FitScale;
        }

        double viewportWidth = Math.Max(1, ImageViewport.ActualWidth);
        double viewportHeight = Math.Max(1, ImageViewport.ActualHeight);
        double widthScale = bitmap.PixelWidth / viewportWidth;
        double heightScale = bitmap.PixelHeight / viewportHeight;
        return Math.Clamp(Math.Max(widthScale, heightScale), FitScale, MaxZoom);
    }
}
