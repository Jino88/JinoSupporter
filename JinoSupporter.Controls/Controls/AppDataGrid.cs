using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 공통 데이터그리드 컨트롤.
    /// </summary>
    public class AppDataGrid : DataGrid
    {
        static AppDataGrid()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(AppDataGrid),
                new FrameworkPropertyMetadata(typeof(AppDataGrid)));
        }
    }
}
