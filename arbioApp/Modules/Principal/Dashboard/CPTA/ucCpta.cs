using arbioApp.Modules.Principal.Dashboard.CIAL;
using arbioApp.Modules.Principal.Dashboard.CIAL.ModulesDash;
using DevExpress.DashboardWin;
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

namespace arbioApp.Modules.Principal.Dashboard.CPTA
{
    public partial class ucCpta : DevExpress.XtraEditors.XtraUserControl
    {
        private DashboardViewer dashboardViewer;
        private static ucCpta _instance;
        private DataTable dt;
        public static ucCpta Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucCpta();
                return _instance;
            }
        }
        public ucCpta()
        {
            InitializeComponent();
            DashboardViewer dashboardViewer = new DashboardViewer();
            dashboardViewer.Dock = DockStyle.Fill;
            string folderPath = Path.Combine(Application.StartupPath, "Dashboard", "Cpta");

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
