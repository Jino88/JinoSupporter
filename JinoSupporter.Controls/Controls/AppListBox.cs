using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 리스트박스 컨트롤.
    /// </summary>
    public class AppListBox : ListBox
    {
        static AppListBox()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppListBox),
                new FrameworkPropertyMetadata(typeof(AppListBox)));
        }
    }
}
