using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using testOutlookAddIn.Properties;

namespace testOutlookAddIn
{
    public partial class ParametersForm : Form
    {
        public ParametersForm()
        {
            InitializeComponent();
        }

        private void ParametersForm_Load(object sender, EventArgs e)
        {
            txtSendUrl.Text = Settings.Default.SendUrl;
            txtBUListUrl.Text = Settings.Default.BUListUrl;
            txtCategoryListUrl.Text = Settings.Default.CategoryListUrl;
            if(listCategory != null)
            {
                foreach (string category in listCategory)
                {
                    cmbCategoryName.Items.Add(category);
                }
                cmbCategoryName.SelectedItem = Settings.Default.CategoryMail;
            }
        }
        public void parmListCategory(List<string> _listCategory)
        {
            listCategory = _listCategory;
        }
        List<string> listCategory;
        private void btnSave_Click(object sender, EventArgs e)
        {
            Settings.Default.SendUrl = txtSendUrl.Text;
            Settings.Default.BUListUrl = txtBUListUrl.Text;
            Settings.Default.CategoryListUrl = txtCategoryListUrl.Text;
            if (cmbCategoryName.SelectedItem != null)
                Settings.Default.CategoryMail = (string)cmbCategoryName.SelectedItem;
            else
                Settings.Default.CategoryMail = "";
            Settings.Default.Save();
        }
    }
}
