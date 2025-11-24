using DevExpress.DashboardCommon;
using DevExpress.DataAccess;
using DevExpress.DataAccess.ConnectionParameters;
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
using DevExpress.DataAccess.Sql;

namespace arbioApp.Modules.Principal.Dashboard.CIAL
{
    public partial class ucDashDesigner : DevExpress.XtraEditors.XtraUserControl
    {
        public ucDashDesigner()
        {
            InitializeComponent();
            dashboardDesigner1.CreateRibbon();
            dashboardDesigner1.CreateCustomItemBars();
            
        }
        public void LoadDashboard(DevExpress.DashboardCommon.Dashboard dashboard)
        {
            dashboardDesigner1.Dashboard = dashboard;

        }
       
    }
}
