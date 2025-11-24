using arbioApp.Modules.Principal.DI._2_Documents;
using DevExpress.DashboardCommon;
using DevExpress.DataAccess;
using DevExpress.DashboardCommon.Native;
using DevExpress.DashboardWin;
using DevExpress.DataAccess.ConnectionParameters;
using DevExpress.DataAccess.Native;
using DevExpress.DataAccess.Sql;
using DevExpress.XtraBars.Navigation;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using arbioApp.Modules.Principal.Dashboard.CIAL.ModulesDash;
using DevExpress.XtraBars;

namespace arbioApp.Modules.Principal.Dashboard.CIAL
{
    public partial class ucCial : DevExpress.XtraEditors.XtraUserControl
    {
        private DashboardViewer dashboardViewer;
        private static ucCial _instance;
        //private DataTable dt;
        public static ucCial Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucCial();
                return _instance;
            }
        }
        private static string connectionString = $"Server=26.53.123.231;Database=arbapp;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=120;";
        
        public ucCial()
        {
            InitializeComponent();

            DashboardViewer dashboardViewer = new DashboardViewer();
            dashboardViewer.Dock = DockStyle.Fill;

            string baseFolderPath = Path.Combine(Application.StartupPath, "Dashboard", "Cial");
            string userEmail = FrmMdiParent.Username;
            string? flag4Raw = null;

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var command = new SqlCommand("SELECT Flag4 FROM T_UserRole WHERE UserName = @Email", connection);
                command.Parameters.AddWithValue("@Email", userEmail);
                flag4Raw = command.ExecuteScalar()?.ToString();
            }

            if (FrmMdiParent.UserRole == "Administrators")
            {
                FillAccordionControlWithXmlFilesGroupedByFolder(baseFolderPath);
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(flag4Raw))
                {
                    // Support multiple depot names separated by comma
                    string[] depots = flag4Raw.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string depot in depots)
                    {
                        string depotFolderPath = Path.Combine(baseFolderPath, depot.Trim());

                        if (Directory.Exists(depotFolderPath))
                        {
                            FillAccordionControlWithXmlFiles(depotFolderPath);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Ajouter un ou plusieurs dépôts dans Flag4 de T_UserRole.");
                }
            }

            hyperlinkLabelControl1.Text = "Designer";
            hyperlinkLabelControl1.Enabled = false;
        }

        private string EquivActivo(string depotArbiochem)
        {
            string depot = depotArbiochem.Trim();
            switch (depot)
            {
                case "AMBOHIMANGAKELY":
                    return "ACTIVO";
                case "ANTANIMORA":
                    return "ACTIVOFEED_ANTANIMORA";
                case "ANALAKELY":
                    return "ACTIVOFEED_ANALAKELY";
                case "DIEGO":
                    return "ACTIVOFEED_DIEGO_AG";
                case "MAHITSY":
                    return "ACTIVOFEED_MAHINTSY";
                case "IMERINTSIATOSIKA":
                    return "ACTIVOFEED_IMERINTSIATOSIKA";
                case "TMM":
                    return "ACTIVOFEED_TMM";
                default:
                    return depot;
            }
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
        private void FillAccordionControlWithXmlFilesGroupedByFolder(string rootFolderPath)
        {
            if (!Directory.Exists(rootFolderPath))
            {
                MessageBox.Show("Le dossier spécifié n'existe pas.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Vider les éléments existants
            accordionControlElement1.Elements.Clear();

            // Parcourir les sous-dossiers
            string[] subDirectories = Directory.GetDirectories(rootFolderPath);

            foreach (string directory in subDirectories)
            {
                // Créer l'élément parent (le dossier)
                AccordionControlElement folderElement = new AccordionControlElement
                {
                    Text = Path.GetFileName(directory),
                    Style = ElementStyle.Group,
                    Expanded = false // Facultatif : pour déplier automatiquement
                };

                // Ajouter les fichiers XML de ce dossier comme enfants
                string[] xmlFiles = Directory.GetFiles(directory, "*.xml");
                foreach (string xmlFile in xmlFiles)
                {
                    AccordionControlElement fileElement = new AccordionControlElement
                    {
                        Text = Path.GetFileNameWithoutExtension(xmlFile),
                        Tag = xmlFile,
                        Style = ElementStyle.Item
                    };

                    folderElement.Elements.Add(fileElement);
                }

                // Ajouter le dossier (et ses fichiers) à l'élément racine
                accordionControlElement1.Elements.Add(folderElement);
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
        //old
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


        private void OpenDesignerForm()
        {
            DevExpress.DashboardCommon.Dashboard dashboard = LoadDashboardFromXml(xmlFilePath);
            ucDashDesigner designerForm = new ucDashDesigner();
            designerForm.Dock = DockStyle.Fill;
            designerForm.LoadDashboard(dashboard);

            panelControl1.Controls.Clear();
            panelControl1.Controls.Add(designerForm);
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


    }
}
