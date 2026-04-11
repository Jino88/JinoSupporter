using System.Windows;
using System.Windows.Controls.Primitives;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 토글버튼 컨트롤.
    /// </summary>
    public class AppToggleButton : ToggleButton
    {
        static AppToggleButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppToggleButton),
                new FrameworkPropertyMetadata(typeof(AppToggleButton)));
        }
    }
}
