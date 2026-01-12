using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmPrix : Form
    {
        public decimal Prix { get; set; }
        public frmPrix(decimal prixActuel)
        {
            InitializeComponent();
            Prix = prixActuel;
        }

        private void btnEnregistrer_Click(object sender, EventArgs e)
        {
            if (!decimal.TryParse(txtpu.Text, out decimal p))
            {
                MessageBox.Show("Prix invalide");
                return;
            }

            if (rdkg.Checked)
            {
                Prix = p;
            }
            else
            {
                Prix = p/1000;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void frmPrix_Load(object sender, EventArgs e)
        {
            txtpu.Text = Prix.ToString("N2");
        }
    }
}
