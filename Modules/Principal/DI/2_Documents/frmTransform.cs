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
    public partial class frmTransform : DevExpress.XtraEditors.XtraForm
    {
       

        public frmTransform(string typdocument)
        {
            InitializeComponent();

            switch (typdocument)
            {
                case "Facture":
                    radioGroup1.Properties.Items[0].Enabled = false;
                    radioGroup1.Properties.Items[1].Enabled = false;
                    radioGroup1.Properties.Items[2].Enabled = false;
                    break;

                case "Bon de livraison":
                    radioGroup1.Properties.Items[0].Enabled = false;
                    radioGroup1.Properties.Items[1].Enabled = false;
                    break;

                case "Bon de commande":
                    radioGroup1.Properties.Items[0].Enabled = false;
                    break;

                case "Bon de réception":
                    radioGroup1.Properties.Items[3].Enabled = false;
                    break;
            }
        }

        public frmEditDocument ParentFormInstance { get; set; }
        public string doctype { get; private set; }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            int selectedIndex = radioGroup1.SelectedIndex;
            doctype = radioGroup1.Properties.Items[selectedIndex].Description;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}