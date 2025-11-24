using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraTreeList.Nodes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace arbioApp.Modules.Principal.GESTION_COMMERCIALE
{
    public partial class FrmNewPrix : DevExpress.XtraEditors.XtraForm
    {
        public static ucChangePrix _ucChangePrix;
        string leserveur;
        string UserRole;
        string UserName;
        public static string selectedSite;
        public static string selectedConnex;
        public FrmNewPrix(ucChangePrix uc)
        {
            InitializeComponent();
            _ucChangePrix = uc;
            leserveur = FrmMdiParent.DataSourceNameValueParent;
            Addserver.LoadLookUpEditSite(leserveur, lookUpEdit1);
            //ExpandAllNodes(
            //
            //EditNewPrix);
        }
        private void ExpandAllNodes(DevExpress.XtraEditors.TreeListLookUpEdit treeList)
        {
            treeList.Properties.TreeList.BeginUpdate();
            try
            {
                foreach (TreeListNode node in treeList.Properties.TreeList.Nodes)
                {
                    ExpandNodeAndChildren(node);
                }
            }
            finally
            {
                treeList.Properties.TreeList.EndUpdate();
            }
        }
        private void ExpandNodeAndChildren(TreeListNode node)
        {
            node.Expanded = true;
            foreach (TreeListNode childNode in node.Nodes)
            {
                ExpandNodeAndChildren(childNode);
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {
                
                selectedSite = lookUpEdit1.EditValue.ToString();
                selectedConnex = FrmMdiParent.DataSourceNameValueParent;
                
                string connectionString;
                if (selectedSite.StartsWith("ACTIVO"))
                {
                    connectionString = @$"Data Source={selectedIp};
                                Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";
                }
                else
                {
                    connectionString = @$"Data Source={selectedConnex};
                                Initial Catalog=ARBIOCHEM;User ID=Dev;Password=1234;TrustServerCertificate=True";
                }
                


                string binDebugPath = Path.Combine(Application.StartupPath);
                try
                {
                    _ucChangePrix.LoadSpreadsheetTemplate(selectedConnex, selectedSite,
                        binDebugPath + "\\ARBIOCHEM.arbpp");
                }
                finally
                {

                }

            }
            catch (System.Exception ex)
            {
                // Gérer l'erreur (afficher un message ou logger)
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Garantir que le Splash Screen est fermé si une autre erreur s'est produite
                SplashScreenManager.CloseForm(false);
            }

            this.Close();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public static string selectedIp;
        private void lookUpEdit1_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEdit editor = sender as LookUpEdit;

            if (editor.EditValue != null)
            {
                DataRowView selectedRow = editor.Properties.GetDataSourceRowByKeyValue(editor.EditValue) as DataRowView;

                if (selectedRow != null)
                {
                    string selectedSite = selectedRow["SiteName"].ToString();
                    selectedIp = selectedRow["address_ip"].ToString();

                    //MessageBox.Show($"Site sélectionné : {selectedSite}\nAdresse IP : {selectedIp}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

    }
}