using FilesEditor.Entities;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace PptGeneratorGUI
{
    public partial class frmSelectFilters : Form
    {
        public InputDataFilters_Item FilterToManage;
        public frmSelectFilters()
        {
            InitializeComponent();

        }
        private void frmSelectFilters_Load(object sender, EventArgs e)
        {
            this.Text = $"Select Filters to Apply to '{FilterToManage.Table} - {FilterToManage.FieldName}'";
            this.lblFieldName.Text = $"{FilterToManage.Table} - {FilterToManage.FieldName}";
            LoadFilterList();
        }

        private void LoadFilterList()
        {
            foreach (var val in FilterToManage.Values.OrderBy(_ => _))
            {
                var itmeChecked = FilterToManage.SelectedValues.Any(_ => _.Equals(val));
                cblFilters.Items.Add(val, itmeChecked);
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            FilterToManage.SelectedValues.Clear();
            foreach (var item in cblFilters.CheckedItems)
            {
                FilterToManage.SelectedValues.Add(item.ToString());
            }
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}