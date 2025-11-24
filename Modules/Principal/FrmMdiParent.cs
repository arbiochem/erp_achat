using DevExpress.LookAndFeel;
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
using arbioApp.Modules.Helpers;
using arbioApp.Modules.Principal.DSI;
using arbioApp.Modules.Principal.Theme;
using arbioApp.Modules.Principal.FINANCE;
using System.Diagnostics;
using arbioApp.Modules.Principal.GESTION_COMMERCIALE;
using static DevExpress.XtraEditors.Mask.Design.MaskSettingsForm.DesignInfo.MaskManagerInfo;
using DevExpress.Utils.VisualEffects;
using DevExpress.XtraBars.Navigation;
using System.Data.SqlClient;
using arbioApp.Modules.Principal;
using DevExpress.XtraSplashScreen;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.Dashboard;

namespace arbioApp
{
    public partial class FrmMdiParent : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
    {
        public static Guid _id_user { get; set; } // uniqueidentifier
        public static string Username { get; set; }
        public static string IDName { get; set; }
        public static string UserRole { get; set; }
        public static string DataSourceNameValueParent { get; set; }
        public static string CatalogNameValueParent { get; set; }
        public FrmMdiParent(Guid id_user, string idname, string username, string userRole, string datasourceName, string catalogName)
        {
            //ThemeHelper.ApplyThemeBasedOnMonth();
            InitializeComponent();
            this.IsMdiContainer = true;
            Username = username;
            UserRole = userRole;
            IDName = idname;
            _id_user = id_user;
            DataSourceNameValueParent = datasourceName;
            CatalogNameValueParent = catalogName;
            ConfigureHyperLinkEdit();

        }
        private async Task CreateContextButtons(AccordionControlElementCollection elements)
        {
            foreach (AccordionControlElement element in elements)
            {
                if (element.Style == ElementStyle.Item)
                {
                    if (element.Text == "Mise à jour des prix")
                    {
                        AccordionContextButton button = new AccordionContextButton
                        {
                            AnimationType = DevExpress.Utils.ContextAnimationType.None,
                            AppearanceNormal = { BackColor = Color.Red },
                            AppearanceHover = { BackColor = Color.LightSalmon },
                            Caption = CountDemande(),
                            Id = Guid.NewGuid(),
                            Padding = new Padding(2),
                            Visibility = DevExpress.Utils.ContextItemVisibility.Visible
                        };
                        element.ContextButtons.Add(button);
                    }
                }

                // Toujours explorer les sous-éléments
                if (element.Elements.Count > 0)
                {
                    await CreateContextButtons(element.Elements);
                }
            }
        }

        private void ConfigureHyperLinkEdit()
        {
            barSubItem1.Caption = IDName;
        }


        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            FrmHelp frmHelp = new FrmHelp();
            frmHelp.Show();
        }

        private async void FrmMdiParent_Load(object sender, EventArgs e)
        {
            var permissionsManager = new PermissionsManager($"Data Source={DataSourceNameValueParent};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True");
            string roleName = UserRole; // Exemple : nom de la colonne du rôle
            var permissions = permissionsManager.LoadPermissions(roleName);

            // Appliquer les permissions au formulaire
            permissionsManager.ApplyPermissions(this, permissions, 1);
            await Task.Delay(10000); 
            if (UserRole == "Administrators")
            {                
                //await CreateContextButtons(accordionControl1.Elements);
            }
        }
        private void VerifierFArtClientDemande()
        {
        }
        private void accordionControlElement2_Click(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(ucReleaseUser.Instance))
            {
                Container.Controls.Add(ucReleaseUser.Instance);
                ucReleaseUser.Instance.Dock = DockStyle.Fill;
                ucReleaseUser.Instance.BringToFront();
            }
            ucReleaseUser.Instance.BringToFront();
        }

        private void barButtonItem1_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FrmProfil frmpr = new FrmProfil();
            frmpr.Show();
            splashScreenManager1.CloseWaitForm();

        }

        private void fluentDesignFormControl1_Click(object sender, EventArgs e)
        {

        }

        private void barButtonTheme_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            FrmTheme frmthm = new FrmTheme();
            frmthm.Show();
        }

        private void accordionControlElement3_Click(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(ucFinanceMvtStock.Instance))
            {
                Container.Controls.Add(ucFinanceMvtStock.Instance);
                ucFinanceMvtStock.Instance.Dock = DockStyle.Fill;
                ucFinanceMvtStock.Instance.BringToFront();
            }
            ucFinanceMvtStock.Instance.BringToFront();
        }

        private void barButtonDisconnect_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (MessageBox.Show("Voulez-vous vraiment vous déconnecter ?", "Déconnexion",
        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
            Process.Start(appPath);
        }
        string appPath = @"C:\arApplication\arbUpdater.exe";

        private void mnuEditPrix_Click(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(Modules.Principal.GESTION_COMMERCIALE.ucChangePrix.Instance))
            {
                Container.Controls.Add(Modules.Principal.GESTION_COMMERCIALE.ucChangePrix.Instance);
                Modules.Principal.GESTION_COMMERCIALE.ucChangePrix.Instance.Dock = DockStyle.Fill;
                Modules.Principal.GESTION_COMMERCIALE.ucChangePrix.Instance.BringToFront();
            }
            Modules.Principal.GESTION_COMMERCIALE.ucChangePrix.Instance.BringToFront();
        }

        public string CountDemande()
        {
            // Récupérer les informations des serveurs et bases de données depuis la table T_Server
            var serveursEtBases = GetServeursEtBases();
            int totalDemandesNonMisesAJour = 0; // Initialisé à 0   
            foreach (var (serveur, baseDeDonnees) in serveursEtBases)
            {
                try
                {
                    string connectionString = $"Server={serveur};Database={baseDeDonnees};User ID=Dev;Password=1234;TrustServerCertificate=True";
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        string query = "SELECT COUNT(UPDATED) FROM F_ARTCLIENT_DEMANDE WHERE UPDATED = 0";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            int count = Convert.ToInt32(cmd.ExecuteScalar());
                            totalDemandesNonMisesAJour += count; // Ajouter au total
                        }
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"Erreur avec le serveur {serveur}: {ex.Message}");
                }
            }

            // Convertir en chaîne après avoir additionné tous les résultats
            return totalDemandesNonMisesAJour.ToString();
        }

        // Méthode pour récupérer les serveurs et bases de données depuis la table T_Server
        private List<(string serveur, string baseDeDonnees)> GetServeursEtBases()
        {
            var serveursEtBases = new List<(string serveur, string baseDeDonnees)>();

            string query = "SELECT Site, connex FROM T_Server"; // La table T_Server contient les informations nécessaires

            // Remplacer par votre chaîne de connexion de base pour récupérer les données
            //string connectionString = "votre_chaine_de_connexion"; 
            string connectionString = $"Data Source={DataSourceNameValueParent};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string serveur = reader["connex"].ToString();
                    string baseDeDonnees = reader["Site"].ToString();

                    serveursEtBases.Add((serveur, baseDeDonnees)); // Ajouter le serveur et la base de données à la liste
                }
            }

            return serveursEtBases;
        }



        private void mnuUpDatePrix_Click_1(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(ucUpdatePrix.Instance))
            {
                Container.Controls.Add(ucUpdatePrix.Instance);
                ucUpdatePrix.Instance.Dock = DockStyle.Fill;
                ucUpdatePrix.Instance.BringToFront();
            }
            ucUpdatePrix.Instance.BringToFront();
        }

        private void accordionControlElement2_Click_1(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(Modules.Principal.GESTION_COMMERCIALE.ucReglement.Instance))
            {
                Container.Controls.Add(Modules.Principal.GESTION_COMMERCIALE.ucReglement.Instance);
                Modules.Principal.GESTION_COMMERCIALE.ucReglement.Instance.Dock = DockStyle.Fill;
                Modules.Principal.GESTION_COMMERCIALE.ucReglement.Instance.BringToFront();
            }
            Modules.Principal.GESTION_COMMERCIALE.ucReglement.Instance.BringToFront();
        }

        
        private void accordionControlElement3_Click_1(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(Modules.Principal.DI._2_Documents.ucDocuments.Instance))
            {
                Container.Controls.Add(Modules.Principal.DI._2_Documents.ucDocuments.Instance);
                Modules.Principal.DI._2_Documents.ucDocuments.Instance.Dock = DockStyle.Fill;
                Modules.Principal.DI._2_Documents.ucDocuments.Instance.BringToFront();
            }
            Modules.Principal.DI._2_Documents.ucDocuments.Instance.BringToFront();
        }

        private void mnuDashCIAL_Click(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(Modules.Principal.Dashboard.CIAL.ucCial.Instance))
            {
                Container.Controls.Add(Modules.Principal.Dashboard.CIAL.ucCial.Instance);
                Modules.Principal.Dashboard.CIAL.ucCial.Instance.Dock = DockStyle.Fill;
                Modules.Principal.Dashboard.CIAL.ucCial.Instance.BringToFront();
            }
            Modules.Principal.Dashboard.CIAL.ucCial.Instance.BringToFront();
        }

        private void mnuDashCPTA_Click(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(Modules.Principal.Dashboard.CPTA.ucCpta.Instance))
            {
                Container.Controls.Add(Modules.Principal.Dashboard.CPTA.ucCpta.Instance);
                Modules.Principal.Dashboard.CPTA.ucCpta.Instance.Dock = DockStyle.Fill;
                Modules.Principal.Dashboard.CPTA.ucCpta.Instance.BringToFront();
            }
            Modules.Principal.Dashboard.CPTA.ucCpta.Instance.BringToFront();
        }


        private void mnuDashRecouvrement_Click(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(Modules.Principal.Dashboard.RECOUVREMENT.ucRecouvrement.Instance))
            {
                Container.Controls.Add(Modules.Principal.Dashboard.RECOUVREMENT.ucRecouvrement.Instance);
                Modules.Principal.Dashboard.RECOUVREMENT.ucRecouvrement.Instance.Dock = DockStyle.Fill;
                Modules.Principal.Dashboard.RECOUVREMENT.ucRecouvrement.Instance.BringToFront();
            }
            Modules.Principal.Dashboard.RECOUVREMENT.ucRecouvrement.Instance.BringToFront();
        }

        private void mnuRecouvClient_Click(object sender, EventArgs e)
        {
            if (!Container.Controls.Contains(Modules.Principal.RECOUVREMENT.ucClient.Instance))
            {
                Container.Controls.Add(Modules.Principal.RECOUVREMENT.ucClient.Instance);
                Modules.Principal.RECOUVREMENT.ucClient.Instance.Dock = DockStyle.Fill;
                Modules.Principal.RECOUVREMENT.ucClient.Instance.BringToFront();
            }
            Modules.Principal.RECOUVREMENT.ucClient.Instance.BringToFront();
        }

        private void Container_Click(object sender, EventArgs e)
        {

        }
    }
}