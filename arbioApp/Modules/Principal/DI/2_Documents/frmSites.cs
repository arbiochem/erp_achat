using DevExpress.DataAccess.Native.Data;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmSites : DevExpress.XtraEditors.XtraForm
    {
        private static ucDocuments _instance;
        private System.Data.DataTable dataTable;
        private SqlDataAdapter dataAdapter;
        private SqlConnection connection;
        public static string connectionString;
        //private frmNewDocument parentForm;
        private frmEditDocument parentEditDocument;

        //public frmSites(frmNewDocument parent)
        //{
        //    InitializeComponent();
        //    connectionString = $"Server={FrmMdiParent.DataSourceNameValueParent};" +
        //                       $"Database=arbapp;User ID=Dev;Password=1234;" +
        //                       $"TrustServerCertificate=True;Connection Timeout=120;";
        //    InitializeGridControlServers();
        //    this.parentForm = parent;
        //}
        public frmSites(frmEditDocument parent)
        {
            InitializeComponent();
            connectionString = $"Server={FrmMdiParent.DataSourceNameValueParent};" +
                               $"Database=arbapp;User ID=Dev;Password=1234;" +
                               $"TrustServerCertificate=True;Connection Timeout=120;";
            InitializeGridControlServers();
            this.parentEditDocument = parent;
        }
        private void InitializeGridControlServers()
        {
            try
            {
                // Définir la requête SQL
                string query = "SELECT ID, actif, BDD, ADDRESS_IP FROM dbo.T_SERVER_DEPOT ORDER BY ADDRESS_IP";

                // Créer la connexion SQL avec la chaîne de connexion
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    // Ouvrir la connexion
                    connection.Open();

                    // Créer un SqlDataAdapter avec la requête SQL et la connexion
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(query, connection);
                    SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);

                    // Créer et remplir le DataTable avec les résultats de la requête
                    dataTable = new System.Data.DataTable();
                    dataAdapter.Fill(dataTable);

                    // Assigner le DataTable comme source de données du GridControl
                    gcServers.DataSource = dataTable;
                    gvServers.PopulateColumns();

                    // Créer un RepositoryItemCheckEdit pour la colonne "actif"
                    RepositoryItemCheckEdit checkEdit = new RepositoryItemCheckEdit()
                    {
                        ValueChecked = true,
                        ValueUnchecked = false,
                        ValueGrayed = false
                    };

                    // Vérifier si la colonne "actif" existe et assigner le RepositoryItemCheckEdit
                    if (gvServers.Columns["actif"] != null)
                    {
                        gvServers.Columns["actif"].ColumnEdit = checkEdit;
                    }

                    // Masquer la colonne "ID"
                    gvServers.Columns["ID"].Visible = false;
                }
            }
            catch (Exception ex)
            {
                // Loguer l'exception complète pour plus de détails
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}\n{ex.StackTrace}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            RefreshGrids();
        }
        private void UpdateEntry(int id, bool actif)
        {
            string query = @"UPDATE dbo.T_SERVER_DEPOT
                     SET actif = @actif
                     WHERE ID = @ID";

            using (var connection = new SqlConnection(connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                connection.Open();
                command.Parameters.AddWithValue("@actif", actif); // bool au lieu de int
                command.Parameters.AddWithValue("@ID", id);

                int rowsAffected = command.ExecuteNonQuery();
                RefreshGrids();
            }
        }
        private void RefreshGrids()
        {
            InitializeGridControlServers();
            gvServers.BestFitColumns();
            gvServers.RefreshData();
            gvServers.InvalidateRows();
        }

        private void gvServers_CellValueChanged(object sender,
            DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName == "actif")
            {
                // Forcer la fin de l'édition
                gvServers.PostEditor();

                // ✅ Récupérer l'ID via RowHandle
                int id = Convert.ToInt32(gvServers.GetRowCellValue(e.RowHandle, "ID"));
                bool actif = Convert.ToBoolean(gvServers.GetRowCellValue(e.RowHandle, "actif"));

                Console.WriteLine($"Mise à jour : ID = {id}, Actif = {actif}");
                UpdateEntry(id, actif);
            }
        }

        
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            
                parentEditDocument.ExecuteStockAlert();
                parentEditDocument.ChargerArtFrns();
                this.Close();
            
        }
        
    }
}