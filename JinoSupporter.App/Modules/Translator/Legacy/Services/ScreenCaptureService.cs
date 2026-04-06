using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using CustomKeyboardCSharp.Windows;

namespace CustomKeyboardCSharp.Services;

public sealed class ScreenCaptureService
{
    public Task<(BitmapSource Image, byte[] ImageBytes)?> CaptureAsync(Window owner)
    {
        var bounds = SystemInformation.VirtualScreen;
        using var screenshot = new Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(screenshot))
        {
            graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, screenshot.Size);
        }

        var selector = new SelectionOverlayWindow(ToBitmapSource(screenshot), bounds)
        {
            Owner = owner
        };

        if (selector.ShowDialog() != true || selector.SelectedRegion is null)
        {
            return null;
        }

        var region = selector.SelectedRegion.Value;
        var cropRectangle = new Rectangle(
            (int)Math.Round(region.X - bounds.Left),
            (int)Math.Round(region.Y - bounds.Top),
            (int)Math.Round(region.Width),
            (int)Math.Round(region.Height));
        using var cropped = screenshot.Clone(cropRectangle, screenshot.PixelFormat);
        var image = ToBitmapSource(cropped);
        return Task.FromResult<(BitmapSource Image, byte[] ImageBytes)?>((
            image,
            ToPngBytes(cropped)
        ));
    }

    private static BitmapSource ToBitmapSource(Bitmap bitmap)
    {
        using var memory = new MemoryStream();
        bitmap.Save(memory, ImageFormat.Png);
        memory.Position = 0;
        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.StreamSource = memory;
        image.EndInit();
        image.Freeze();
        return image;
    }

    private static byte[] ToPngBytes(Bitmap bitmap)
    {
        using var memory = new MemoryStream();
        bitmap.Save(memory, ImageFormat.Png);
        return memory.ToArray();
    }
}
