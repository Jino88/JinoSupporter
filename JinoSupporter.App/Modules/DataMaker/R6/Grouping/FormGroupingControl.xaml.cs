using DataMaker.R6.Grouping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace DataMaker.R6.FormGrouping
{
    public partial class FormGroupingControl : UserControl
    {
        public int index = 0;

        public FormGroupingControl(int index)
        {
            InitializeComponent();
            this.index = index;
        }

        public void SetModelData(List<string> Models)
        {
            CT_LIST.Items.Clear();
            foreach (var model in Models)
            {
                CT_LIST.Items.Add(model);
            }
        }

        public void SetGroupName(string groupName)
        {
            if (CT_TB_GROUPNAME != null)
            {
                CT_TB_GROUPNAME.Text = groupName ?? "";
            }
        }

        public clModelGroupData GetSelectedModels()
        {
            List<string> selectedModels = new List<string>();
            foreach (var item in CT_LIST.SelectedItems)
            {
                selectedModels.Add(item.ToString());
            }

            string groupName = CT_TB_GROUPNAME.Text.Trim();
            if (string.IsNullOrEmpty(groupName))
            {
                groupName = string.Join("_", selectedModels.OrderBy(m => m));
            }

            clModelGroupData data = new clModelGroupData()
            {
                Index = index,
                ModelList = selectedModels,
                GroupName = groupName
            };

            return data;
        }

        public void SetCheckinGroup(clModelGroupData groupData)
        {
            foreach (var model in groupData.ModelList)
            {
                for (int i = 0; i < CT_LIST.Items.Count; i++)
                {
                    if (CT_LIST.Items[i].ToString() == model)
                    {
                        CT_LIST.SelectedItems.Add(CT_LIST.Items[i]);
                        break;
                    }
                }
            }

            SetGroupName(groupData.GroupName);
        }
    }
}
