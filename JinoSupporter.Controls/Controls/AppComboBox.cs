using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 콤보박스 컨트롤.
    /// </summary>
    public class AppComboBox : ComboBox
    {
        static AppComboBox()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppComboBox),
                new FrameworkPropertyMetadata(typeof(AppComboBox)));
        }
    }
}
