using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 프로그레스바 컨트롤.
    /// </summary>
    public class AppProgressBar : ProgressBar
    {
        static AppProgressBar()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppProgressBar),
                new FrameworkPropertyMetadata(typeof(AppProgressBar)));
        }
    }
}
