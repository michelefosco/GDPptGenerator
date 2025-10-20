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
            AggiornaTitolo();
            LoadFilterList();
        }

        private void AggiornaTitolo()
        {
            this.Text = $"Select Filters to Apply to '{FilterToManage.Table} - {FilterToManage.FieldName}'";
            this.lblFieldName.Text = $"{FilterToManage.Table} - {FilterToManage.FieldName}";
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

            // se tutti i valori sono selezionati è equivalente ad non averne, quindi lascio la lista vuota
            if (cblFilters.CheckedItems.Count != cblFilters.Items.Count)
            {
                // collego gli elementi selezionati ai valori del filtro
                foreach (var item in cblFilters.CheckedItems)
                { FilterToManage.SelectedValues.Add(item.ToString()); }
            }

            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bntSelectAll_Click(object sender, EventArgs e)
        {
            cblFilters_AllItmesChecked(true);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            cblFilters_AllItmesChecked(false);
        }

        private void cblFilters_AllItmesChecked(bool check)
        {
            for (int i = 0; i < cblFilters.Items.Count; i++)
            { cblFilters.SetItemChecked(i, check); }
        }
    }
}