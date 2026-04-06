using System;
using System.Windows;

namespace DataMaker
{
    public partial class ProgressPopupWindow : Window
    {
        public ProgressPopupWindow()
        {
            InitializeComponent();
        }

        public void UpdateState(string title, string message, int taskProgress, int totalProgress)
        {
            int safeTask = Math.Clamp(taskProgress, 0, 100);
            int safeTotal = Math.Clamp(totalProgress, 0, 100);

            CT_TB_TITLE.Text = string.IsNullOrWhiteSpace(title) ? "Processing" : title;
            CT_TB_MESSAGE.Text = string.IsNullOrWhiteSpace(message) ? "Please wait..." : message;
            CT_PROG_TASK.Value = safeTask;
            CT_PROG_TOTAL.Value = safeTotal;
            CT_TB_TASK_PERCENT.Text = $"{safeTask}%";
            CT_TB_TOTAL_PERCENT.Text = $"{safeTotal}%";
            CT_TB_PERCENT.Text = $"Current Task {safeTask}% | Total Task {safeTotal}%";
        }

        public void CenterToOwnerOrScreen()
        {
            Window? target = Owner;
            if (target is { IsLoaded: true, WindowState: not WindowState.Minimized })
            {
                double ownerLeft = target.Left;
                double ownerTop = target.Top;
                double ownerWidth = target.ActualWidth > 0 ? target.ActualWidth : target.Width;
                double ownerHeight = target.ActualHeight > 0 ? target.ActualHeight : target.Height;

                Left = ownerLeft + Math.Max(0, (ownerWidth - Width) / 2);
                Top = ownerTop + Math.Max(0, (ownerHeight - Height) / 2);
                return;
            }

            Rect workArea = SystemParameters.WorkArea;
            Left = workArea.Left + Math.Max(0, (workArea.Width - Width) / 2);
            Top = workArea.Top + Math.Max(0, (workArea.Height - Height) / 2);
        }
    }
}
