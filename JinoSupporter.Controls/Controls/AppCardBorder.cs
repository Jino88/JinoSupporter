using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 카드 패널 컨트롤.
    /// 모듈 내 섹션/카드 구분에 사용.
    /// </summary>
    public class AppCardBorder : Border
    {
        static AppCardBorder()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppCardBorder),
                new FrameworkPropertyMetadata(typeof(AppCardBorder)));
        }
    }
}
