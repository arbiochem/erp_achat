using arbioApp.Modules.Principal.Dashboard.CIAL;
using arbioApp.Modules.Principal.Dashboard.CIAL.ModulesDash;
using DevExpress.DashboardCommon.Native;
using DevExpress.DashboardCommon;
using DevExpress.DashboardWin;
using DevExpress.DataAccess.Sql;
using DevExpress.XtraBars.Navigation;
using DevExpress.XtraEditors;
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
using DevExpress.DataAccess;

namespace arbioApp.Modules.Principal.Dashboard.RECOUVREMENT
{
    public partial class ucRecouvrement : DevExpress.XtraEditors.XtraUserControl
    {
        private DashboardViewer dashboardViewer;
        private static ucRecouvrement _instance;
        private DataTable dt;
        public static ucRecouvrement Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucRecouvrement();
                return _instance;
            }
        }
        public ucRecouvrement()
        {
            InitializeComponent();
            DashboardViewer dashboardViewer = new DashboardViewer();
            dashboardViewer.Dock = DockStyle.Fill;
            string folderPath = Path.Combine(Application.StartupPath, "Dashboard", "RECOUVREMENT");

            FillAccordionControlWithXmlFiles(folderPath);

            hyperlinkLabelControl1.Text = "Designer";
            hyperlinkLabelControl1.Enabled = false;
            
        }
        private void FillAccordionControlWithXmlFiles(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show("Le dossier spécifié n'existe pas.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string[] xmlFiles = Directory.GetFiles(folderPath, "*.xml");

            foreach (var xmlFile in xmlFiles)
            {
                AccordionControlElement element = new AccordionControlElement
                {
                    Text = Path.GetFileName(xmlFile.Replace(".xml", "")),
                    Tag = xmlFile
                };

                accordionControlElement1.Elements.Add(element);
            }
        }
        private string xmlFilePath = "";
        private void accordionControl1_ElementClick(object sender, ElementClickEventArgs e)
        {
            try
            {
                if (e.Element.Elements.Count > 0)
                {
                    return;
                }

                if (e.Element is AccordionControlElement element)
                {
                    xmlFilePath = e.Element.Tag.ToString();
                    LoadDashboardInViewer(xmlFilePath);
                    if (FrmMdiParent.UserRole == "Administrators")
                    {
                        hyperlinkLabelControl1.Enabled = true;
                    }
                    else
                    {
                        hyperlinkLabelControl1.Enabled = false;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void LoadDashboardInViewer(string xmlFilePath)
        {
            DevExpress.DashboardCommon.Dashboard dashboard = LoadDashboardFromXml(xmlFilePath);
            
            ucdashDocVente viewerUserControl = new ucdashDocVente();
            viewerUserControl.Dock = DockStyle.Fill;
            viewerUserControl.LoadDashboard(dashboard);
            panelControl1.Controls.Clear();
            panelControl1.Controls.Add(viewerUserControl);
        }
        //private void LoadDashboardInViewer(string xmlFilePath)
        //{
        //    DevExpress.DashboardCommon.Dashboard dashboard = LoadDashboardFromXml(xmlFilePath);

        //    // Correction de l'erreur CS7036 : Ajout du paramètre 'lookupsettings'
        //    StaticListLookUpSettings staticSettings = new StaticListLookUpSettings();
        //    //staticSettings.Values = new object[] { DateTime.Now.AddYears(-1), DateTime.Now, DateTime.Now.AddYears(1) };

        //    DashboardParameter dateParameter = new DashboardParameter(
        //        "dateParameter",
        //        typeof(DateTime),
        //        DateTime.Now,  // valeur par défaut
        //        "Sélectionnez la date :",
        //        true,
        //        staticSettings // Paramètre ajouté pour résoudre l'erreur
        //    );
        //    dashboard.Parameters.Add(dateParameter);

        //    // Configuration de la source de données
        //    DashboardSqlDataSource dataSource = (DashboardSqlDataSource)dashboard.DataSources[0];
        //    CustomSqlQuery salesPersonQuery = (CustomSqlQuery)dataSource.Queries[0];

        //    // Ajout du paramètre à la requête SQL
        //    salesPersonQuery.Parameters.Add(new QueryParameter("startDate", typeof(Expression),
        //        new Expression("[Parameters.dateParameter]")));

        //    // Requête SQL avec paramètre
        //    salesPersonQuery.Sql =
        //        "SELECT * FROM VW_IMPAYE_VALIDE WHERE startdate = getdate()";//@startDate

        //    // Chargement du dashboard
        //    ucdashDocVente viewerUserControl = new ucdashDocVente();
        //    viewerUserControl.Dock = DockStyle.Fill;
        //    viewerUserControl.LoadDashboard(dashboard);
        //    panelControl1.Controls.Clear();
        //    panelControl1.Controls.Add(viewerUserControl);
        //}

        private DevExpress.DashboardCommon.Dashboard LoadDashboardFromXml(string xmlFilePath)
        {
            DevExpress.DashboardCommon.Dashboard dashboard = new DevExpress.DashboardCommon.Dashboard();
            dashboard.LoadFromXml(xmlFilePath);
            return dashboard;
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            if (hyperlinkLabelControl1.Text == "Designer")
            {
                hyperlinkLabelControl1.Text = "Viewer";
                OpenDesignerForm();

            }
            else
            {
                hyperlinkLabelControl1.Text = "Designer";
                LoadDashboardInViewer(xmlFilePath);
            }
        }
        private void OpenDesignerForm()
        {
            DevExpress.DashboardCommon.Dashboard dashboard = LoadDashboardFromXml(xmlFilePath);
            ucDashDesigner designerForm = new ucDashDesigner();
            designerForm.Dock = DockStyle.Fill;
            designerForm.LoadDashboard(dashboard);

            panelControl1.Controls.Clear();
            panelControl1.Controls.Add(designerForm);
        }
    }
}
