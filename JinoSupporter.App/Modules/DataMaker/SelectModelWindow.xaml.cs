using System.Collections.Generic;
using System.Windows;

namespace DataMaker
{
    public partial class SelectModelWindow : Window
    {
        public int index = 0;

        public List<string> ModelNames { get; private set; }

        public SelectModelWindow(int idx, List<string> ModelNames)
        {
            InitializeComponent();
            index = idx;
            this.ModelNames = ModelNames;
            CT_LB.Text = index.ToString();
            Init();
        }

        public void Init()
        {
            CT_CB_MODEL.Items.Clear();
            foreach (var model in ModelNames)
            {
                CT_CB_MODEL.Items.Add(model);
            }
            if (CT_CB_MODEL.Items.Count > 0)
            {
                CT_CB_MODEL.SelectedIndex = 0;
            }
        }

        public void SetSelection(clSelectOption opt)
        {
            CT_LB.Text = index.ToString();
            CT_CB_MODEL.SelectedItem = opt.SelectModel;
        }

        public string GetSelectedModel()
        {
            string selectedModel = CT_CB_MODEL.SelectedItem?.ToString();
            if (selectedModel == null)
            {
                return null;
            }
            return selectedModel;
        }
    }
}
