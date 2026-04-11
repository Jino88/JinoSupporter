using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 체크박스 컨트롤.
    /// </summary>
    public class AppCheckBox : CheckBox
    {
        static AppCheckBox()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppCheckBox),
                new FrameworkPropertyMetadata(typeof(AppCheckBox)));
        }
    }
}
