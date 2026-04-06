using System;
using System.Drawing;
using System.Windows;
using System.Windows.Input;

namespace WorkbenchHost.Modules.ScreenCapture;

public partial class RegionCaptureOverlayWindow : Window
{
    private System.Windows.Point? _startPoint;

    public Rectangle SelectedBounds { get; private set; } = Rectangle.Empty;

    public RegionCaptureOverlayWindow()
    {
        InitializeComponent();

        Left = SystemParameters.VirtualScreenLeft;
        Top = SystemParameters.VirtualScreenTop;
        Width = SystemParameters.VirtualScreenWidth;
        Height = SystemParameters.VirtualScreenHeight;
    }

    private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _startPoint = e.GetPosition(this);
        SelectionBorder.Visibility = Visibility.Visible;
        System.Windows.Controls.Canvas.SetLeft(SelectionBorder, _startPoint.Value.X);
        System.Windows.Controls.Canvas.SetTop(SelectionBorder, _startPoint.Value.Y);
        SelectionBorder.Width = 0;
        SelectionBorder.Height = 0;
        CaptureMouse();
    }

    private void Window_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (_startPoint is not System.Windows.Point startPoint || e.LeftButton != MouseButtonState.Pressed)
        {
            return;
        }

        System.Windows.Point current = e.GetPosition(this);
        UpdateSelectionVisual(startPoint, current);
    }

    private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_startPoint is not System.Windows.Point startPoint)
        {
            return;
        }

        System.Windows.Point endPoint = e.GetPosition(this);
        ReleaseMouseCapture();
        SelectedBounds = ToRectangle(startPoint, endPoint);
        DialogResult = SelectedBounds.Width > 1 && SelectedBounds.Height > 1;
        Close();
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape)
        {
            return;
        }

        DialogResult = false;
        Close();
    }

    private void UpdateSelectionVisual(System.Windows.Point start, System.Windows.Point end)
    {
        double left = Math.Min(start.X, end.X);
        double top = Math.Min(start.Y, end.Y);
        double width = Math.Abs(end.X - start.X);
        double height = Math.Abs(end.Y - start.Y);

        System.Windows.Controls.Canvas.SetLeft(SelectionBorder, left);
        System.Windows.Controls.Canvas.SetTop(SelectionBorder, top);
        SelectionBorder.Width = width;
        SelectionBorder.Height = height;
    }

    private Rectangle ToRectangle(System.Windows.Point start, System.Windows.Point end)
    {
        int left = (int)Math.Round(Math.Min(start.X, end.X) + Left);
        int top = (int)Math.Round(Math.Min(start.Y, end.Y) + Top);
        int right = (int)Math.Round(Math.Max(start.X, end.X) + Left);
        int bottom = (int)Math.Round(Math.Max(start.Y, end.Y) + Top);
        return Rectangle.FromLTRB(left, top, right, bottom);
    }
}
