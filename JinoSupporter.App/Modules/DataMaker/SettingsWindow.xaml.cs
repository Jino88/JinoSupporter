using DataMaker.R6;
using System.Windows;

namespace DataMaker
{
    public partial class SettingsWindow : Window
    {
        public string RoutingPath { get; set; }
        public string ReasonPath { get; set; }

        public SettingsWindow(string routingPath, string reasonPath)
        {
            InitializeComponent();

            RoutingPath = routingPath;
            ReasonPath = reasonPath;

            LoadValues();
        }

        private void LoadValues()
        {
            TB_RoutingPath.Text = RoutingPath;
            TB_ReasonPath.Text = ReasonPath;

            TB_TopNGCount.Text = CONSTANT.OPTION.TopNGCount.ToString();
            TB_TopProcessCount.Text = CONSTANT.OPTION.TopProcessCount.ToString();
            TB_TopDefectsPerReason.Text = CONSTANT.OPTION.TopDefectsPerReason.ToString();
            TB_WeekOffset.Text = CONSTANT.OPTION.WorstRankingWeekOffset.ToString();
            TB_MonthOffset.Text = CONSTANT.OPTION.WorstRankingMonthOffset.ToString();
            TB_nQUERY.Text = CONSTANT.OPTION.nQUERY.ToString();
            TB_nQtyWorst.Text = CONSTANT.OPTION.nQtyWorst.ToString();

            CB_RankingPeriodType.SelectedIndex = CONSTANT.OPTION.RankingPeriodType == "Month" ? 1 : 0;
        }

        private void CT_BT_SAVE_Click(object sender, RoutedEventArgs e)
        {
            RoutingPath = TB_RoutingPath.Text.Trim();
            ReasonPath = TB_ReasonPath.Text.Trim();

            if (int.TryParse(TB_TopNGCount.Text, out int topNG)) CONSTANT.OPTION.TopNGCount = topNG;
            if (int.TryParse(TB_TopProcessCount.Text, out int topProc)) CONSTANT.OPTION.TopProcessCount = topProc;
            if (int.TryParse(TB_TopDefectsPerReason.Text, out int topDef)) CONSTANT.OPTION.TopDefectsPerReason = topDef;
            if (int.TryParse(TB_WeekOffset.Text, out int weekOff)) CONSTANT.OPTION.WorstRankingWeekOffset = weekOff;
            if (int.TryParse(TB_MonthOffset.Text, out int monthOff)) CONSTANT.OPTION.WorstRankingMonthOffset = monthOff;
            if (int.TryParse(TB_nQUERY.Text, out int nQuery)) CONSTANT.OPTION.nQUERY = nQuery;
            if (int.TryParse(TB_nQtyWorst.Text, out int nQty)) CONSTANT.OPTION.nQtyWorst = nQty;

            CONSTANT.OPTION.RankingPeriodType = CB_RankingPeriodType.SelectedIndex == 1 ? "Month" : "Week";

            DialogResult = true;
            Close();
        }

        private void CT_BT_CANCEL_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
