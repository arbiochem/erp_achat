using DevExpress.Xpo;
using DevExpress.XtraEditors;
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
    public partial class frmQuery : DevExpress.XtraEditors.XtraForm
    {
        private string _query;

        public frmQuery(string query)
        {
            InitializeComponent();
            _query = query;
            DisplayQuery();
        }
        private void DisplayQuery()
        {
            // Afficher la requête dans une TextBox
            memoEdit1.Text = _query;
            memoEdit1.SelectionStart = 0;
            memoEdit1.SelectionLength = 0;

            // Optionnel : coloration syntaxique basique
            memoEdit1.Font = new Font("Consolas", 10);
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(_query);
            MessageBox.Show("Requête copiée dans le presse-papier !", "Succès",
                           MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}