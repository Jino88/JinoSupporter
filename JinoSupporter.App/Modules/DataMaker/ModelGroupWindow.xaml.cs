using DataMaker.Logger;
using DataMaker.R6.FormGrouping;
using DataMaker.R6.Grouping;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace DataMaker
{
    public partial class ModelGroupWindow : Window
    {
        private List<FormGroupingControl> ListFormGrouping = new List<FormGroupingControl>();
        private List<string> ListUniqueModelName;
        private int idxFormGrouping = 0;

        public List<clModelGroupData> ResultModelGroupData { get; private set; }

        public ModelGroupWindow(List<string> uniqueModelNames)
        {
            InitializeComponent();
            ListUniqueModelName = uniqueModelNames;

            if (ListUniqueModelName != null && ListUniqueModelName.Count > 0)
            {
                AddGroupPanel();
            }
        }

        private void AddGroupPanel()
        {
            var f = new FormGroupingControl(idxFormGrouping++);
            f.SetModelData(ListUniqueModelName);
            ListFormGrouping.Add(f);
            CT_PANEL_MODEL.Children.Add(f);
        }

        private void CT_BT_PLUS_Click(object sender, RoutedEventArgs e)
        {
            if (ListUniqueModelName == null || ListUniqueModelName.Count == 0)
            {
                MessageBox.Show("No model data loaded.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            AddGroupPanel();
        }

        private void CT_BT_MINUS_Click(object sender, RoutedEventArgs e)
        {
            if (CT_PANEL_MODEL.Children.Count > 0)
            {
                CT_PANEL_MODEL.Children.RemoveAt(CT_PANEL_MODEL.Children.Count - 1);
                ListFormGrouping.RemoveAt(ListFormGrouping.Count - 1);
            }
        }

        private void CT_BT_SAVE_Click(object sender, RoutedEventArgs e)
        {
            var data = GetModelGroupData();
            if (data.Count == 0)
            {
                MessageBox.Show("No model groups to save.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "JSON Files|*.json",
                Title = "Save Model Groups"
            };
            if (dlg.ShowDialog() != true) return;

            MainWindow.SaveToJson(data, dlg.FileName);
        }

        private void CT_BT_LOAD_Click(object sender, RoutedEventArgs e)
        {
            if (ListUniqueModelName == null || ListUniqueModelName.Count == 0)
            {
                MessageBox.Show("No model data available.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new OpenFileDialog
            {
                Filter = "JSON Files|*.json",
                Title = "Load Model Groups"
            };
            if (dlg.ShowDialog() != true) return;

            var loaded = MainWindow.LoadFromJson(dlg.FileName);

            CT_PANEL_MODEL.Children.Clear();
            ListFormGrouping = new List<FormGroupingControl>();
            idxFormGrouping = 0;

            foreach (var modelData in loaded)
            {
                var f = new FormGroupingControl(idxFormGrouping++);
                f.SetModelData(ListUniqueModelName);
                f.SetCheckinGroup(modelData);
                if (!string.IsNullOrEmpty(modelData.GroupName))
                    f.SetGroupName(modelData.GroupName);
                ListFormGrouping.Add(f);
                CT_PANEL_MODEL.Children.Add(f);
            }
        }

        private List<clModelGroupData> GetModelGroupData()
        {
            var result = new List<clModelGroupData>();
            foreach (var f in ListFormGrouping)
            {
                var selected = f.GetSelectedModels();
                if (selected.ModelList.Count > 0)
                    result.Add(selected);
            }
            return result;
        }
    }
}
