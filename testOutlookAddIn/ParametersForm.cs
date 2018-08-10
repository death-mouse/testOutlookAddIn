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
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Settings.Default.SendUrl = txtSendUrl.Text;
            Settings.Default.BUListUrl = txtBUListUrl.Text;
            Settings.Default.CategoryListUrl = txtCategoryListUrl.Text;

            Settings.Default.Save();
        }
    }
}
