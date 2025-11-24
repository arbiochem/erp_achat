using DevExpress.DashboardWin;
using DevExpress.XtraEditors;
using DevExpress.XtraPivotGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.Dashboard.CIAL.ModulesDash
{
    public partial class ucdashDocVente : DevExpress.XtraEditors.XtraUserControl
    {
        public ucdashDocVente()
        {
            InitializeComponent();
        }
       
        public void LoadDashboard(DevExpress.DashboardCommon.Dashboard dashboard)
        {
            dashboardViewer1.Dashboard = dashboard;
        }

       
    }
}
