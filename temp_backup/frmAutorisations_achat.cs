using arbioApp.Models;
using DevExpress.XtraCharts.Designer.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
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
    public partial class frmAutorisations_achat : DevExpress.XtraEditors.XtraForm
    {
        private SqlDataAdapter adapter;
        private DataTable dt;
        private SqlCommandBuilder builder;
        DataTable dtUsers;
        public frmAutorisations_achat()
        {
            InitializeComponent();
        }

        private void frmAutorisations_achat_Load(object sender, EventArgs e)
        {
            LoadUsers();
            LoadAutorisation();

        }
        private void LoadUsers()
        {
            string connectionString = $"Server=SRV-ARB;" +
                                          $"Database=arbapp;User ID=Dev;Password=1234;" +
                                          $"TrustServerCertificate=True;Connection Timeout=120;";
            //string query = "SELECT Username FROM dbo.T_UserRole WHERE UserGroup like '%Achat%'";
            string query = "SELECT Username FROM dbo.T_UserRole";

            using (var da = new SqlDataAdapter(query, connectionString))
            {
                dtUsers = new DataTable();
                da.Fill(dtUsers);
            }
        }
        private void LoadAutorisation()
        {
            try
            {
                gd_autorisation_achat.DataSource = null;
                gv_achat.Columns.Clear();
                string connectionString = $"Server=SRV-ARB;" +
                                          $"Database=arbapp;User ID=Dev;Password=1234;" +
                                          $"TrustServerCertificate=True;Connection Timeout=120;";
                string query = "SELECT * FROM dbo.AUTORISATIONS_ACHAT";

                adapter = new SqlDataAdapter(query, connectionString);
                builder = new SqlCommandBuilder(adapter); // auto-génère INSERT, UPDATE, DELETE

                dt = new DataTable();
                try
                {
                    adapter.Fill(dt);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur lors du remplissage : " + ex.Message);
                }
                gd_autorisation_achat.DataSource = dt;

                // Configurer le GridView
                var view = gv_achat;
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
            repoLookup.DisplayMember = "Username";
            repoLookup.ValueMember = "Username"; 
            repoLookup.NullText = "";
            repoLookup.Columns.Clear();
            repoLookup.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Username", "Username"));
            repoLookup.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repoLookup.SearchMode = DevExpress.XtraEditors.Controls.SearchMode.AutoComplete;
            gv_achat.PopulateColumns();

            // Associer à la colonne mailUser
            gv_achat.Columns["mailUser"].ColumnEdit = repoLookup;

            var typeDocuments = new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>("PA", "Projet d'achat"),
                new KeyValuePair<string, string>("BC", "Bon de commande"),
                new KeyValuePair<string, string>("FA", "Facture"),
                new KeyValuePair<string, string>("PL", "Packing list"),
                new KeyValuePair<string, string>("BR", "Bon de réception")
            };

            var repoTypeDoc = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
            repoTypeDoc.DataSource = typeDocuments;
            repoTypeDoc.DisplayMember = "Value"; // Ce qui est affiché dans la liste
            repoTypeDoc.ValueMember = "Value";     // Ce qui est stocké dans la cellule
            repoTypeDoc.NullText = "";

            repoTypeDoc.Columns.Clear();
            repoTypeDoc.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Type de document"));

            // Optionnel : auto-complétion + saisie directe désactivée
            repoTypeDoc.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repoTypeDoc.SearchMode = DevExpress.XtraEditors.Controls.SearchMode.AutoComplete;

            // 3. Lier l’éditeur à la colonne "type_document"
            gv_achat.Columns["type_document"].ColumnEdit = repoTypeDoc;

            // 4. Ne pas oublier d'ajouter le repository à ta grille
            gd_autorisation_achat.RepositoryItems.Add(repoTypeDoc);
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            gv_achat.CloseEditor();
            gv_achat.UpdateCurrentRow();

            try
            {
                adapter.Update(dt);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frmAutorisations_achat_Load(sender, e);

            MessageBox.Show("Autorisation effectuée pour cet utilisateur", "Affectation de droits", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void frmAutorisations_achat_Activated(object sender, EventArgs e)
        {
            gv_achat.OptionsView.NewItemRowPosition = NewItemRowPosition.Bottom;
            gv_achat.ValidateRow += (s, e) =>
            {
                GridView view = s as GridView;
                DataRow row = view.GetDataRow(e.RowHandle);

                if (row != null)
                {
                    // Exemple : vérifier qu'un champ obligatoire n'est pas vide
                    if (row["mailUser"] == DBNull.Value || string.IsNullOrWhiteSpace(row["mailUser"].ToString()))
                    {
                        e.Valid = false;
                        view.SetColumnError(view.Columns["mailUser"], "Le nom d'utilisateur est obligatoire.");
                    }
                }
            };
        }
    }
}