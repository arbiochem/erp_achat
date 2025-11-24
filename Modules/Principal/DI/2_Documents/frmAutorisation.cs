using arbioApp.Modules.Helpers;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraSplashScreen;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmAutorisation : DevExpress.XtraEditors.XtraForm
    {
        private SqlDataAdapter adapter;
        private DataTable dt;
        private SqlCommandBuilder builder;
        DataTable dtUsers;
        public frmAutorisation()
        {
            InitializeComponent();
        }

        private void frmAutorisation_Load(object sender, EventArgs e)
        {
            LoadUsers();
            LoadAutorisiation();
            
        }
        private void LoadUsers()
        {
            string connectionString = $"Server=SRV-ARB;" +
                          $"Database=arbapp;User ID=Dev;Password=1234;" +
                          $"TrustServerCertificate=True;Connection Timeout=120;";
            string query = "SELECT Username FROM dbo.T_UserRole";

            using (var da = new SqlDataAdapter(query, connectionString))
            {
                dtUsers = new DataTable();
                da.Fill(dtUsers);
            }
        }
        private void LoadAutorisiation()
        {
            try
            {
                gcAut.DataSource = null;
                gvAut.Columns.Clear();
                string connectionString = $"Server=SRV-ARB;" +
                                          $"Database=arbapp;User ID=Dev;Password=1234;" +
                                          $"TrustServerCertificate=True;Connection Timeout=120;";
                string query = "SELECT * FROM dbo.DI_AUTORISATIONS";

                adapter = new SqlDataAdapter(query, connectionString);
                builder = new SqlCommandBuilder(adapter); // auto-génère INSERT, UPDATE, DELETE

                dt = new DataTable();
                adapter.Fill(dt);

                gcAut.DataSource = dt;

                // Configurer le GridView
                var view = gvAut;
                view.OptionsBehavior.Editable = true;
                SetupMailUserLookup();

            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
        private void SetupMailUserLookup()
        {
            var repoLookup = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();

            repoLookup.DataSource = dtUsers;
            repoLookup.DisplayMember = "userName"; // affiché dans la combo
            repoLookup.ValueMember = "userName";   // valeur stockée dans mailUser
            repoLookup.NullText = "";              // si pas de valeur
            repoLookup.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repoLookup.SearchMode = DevExpress.XtraEditors.Controls.SearchMode.AutoComplete;
            
            // Associer à la colonne mailUser
            gvAut.Columns["mailUser"].ColumnEdit = repoLookup;
           
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            gvAut.CloseEditor();
            gvAut.UpdateCurrentRow();
            adapter.Update(dt);
            MessageBox.Show("Autorisation effectuée pour cet utilisateur", "Affectation de droits", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}