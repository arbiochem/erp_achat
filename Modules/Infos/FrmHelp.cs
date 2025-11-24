using DevExpress.XtraEditors;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp
{
    public partial class FrmHelp : DevExpress.XtraEditors.XtraForm
    {
        public FrmHelp()
        {
            InitializeComponent();
            labelControl1.Text = "v1.0.0";
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}