using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 버튼 컨트롤.
    /// ButtonType 속성으로 Default / Accent / Danger / Success / Ghost 변형을 지정.
    /// </summary>
    public class AppButton : Button
    {
        public static readonly DependencyProperty ButtonTypeProperty =
            DependencyProperty.Register(
                nameof(ButtonType),
                typeof(AppButtonType),
                typeof(AppButton),
                new PropertyMetadata(AppButtonType.Default));

        public AppButtonType ButtonType
        {
            get => (AppButtonType)GetValue(ButtonTypeProperty);
            set => SetValue(ButtonTypeProperty, value);
        }

        static AppButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppButton),
                new FrameworkPropertyMetadata(typeof(AppButton)));
        }
    }
}
