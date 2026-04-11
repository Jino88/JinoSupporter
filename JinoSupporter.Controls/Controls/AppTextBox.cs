using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 텍스트박스 컨트롤.
    /// Placeholder 속성으로 힌트 텍스트를 표시.
    /// </summary>
    public class AppTextBox : TextBox
    {
        public static readonly DependencyProperty PlaceholderProperty =
            DependencyProperty.Register(
                nameof(Placeholder),
                typeof(string),
                typeof(AppTextBox),
                new PropertyMetadata(string.Empty));

        public string Placeholder
        {
            get => (string)GetValue(PlaceholderProperty);
            set => SetValue(PlaceholderProperty, value);
        }

        static AppTextBox()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppTextBox),
                new FrameworkPropertyMetadata(typeof(AppTextBox)));
        }
    }
}
