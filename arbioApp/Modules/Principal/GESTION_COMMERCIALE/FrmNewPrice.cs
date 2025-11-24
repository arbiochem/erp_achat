using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using System;
using System.Data;
using System.Data.SqlClient;

namespace arbioApp.Modules.Principal.GESTION_COMMERCIALE
{
    public partial class FrmNewPrice : DevExpress.XtraEditors.XtraForm
    {
        public static string _datasource;
        public static string _catalog;
        public static string _arRef;
        public static string _designation;
        public static string _categorie;
        public static decimal _prixActuel;
        public static  decimal _newPrice;
        public static DataTable myDataTable;

        public static ucChangePrix _ucChangePrix;
        public FrmNewPrice(ucChangePrix uc,  string arRef, string designation, string categorie, decimal prixActuel, decimal nouveauprix)
        {
            InitializeComponent();
            _ucChangePrix = uc;
            _arRef = arRef;
            _designation = designation;
            _categorie = categorie;
            _prixActuel = prixActuel;

            // Exemple d'affichage dans des TextBox
            txtRef.Text = _arRef;
            txtDesign.Text = _designation;
            txtCatTar.Text = _categorie;
            txtPrixActuel.EditValue = prixActuel;
            txtPrixActuel.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            txtPrixActuel.Properties.Mask.EditMask = "n2";
            
        }

        private void FrmNewPrice_Shown(object sender, EventArgs e)
        {
            txtNewPrice.Focus();
        }
       

        private void FrmNewPrice_Load(object sender, EventArgs e)
        {
            string connectionString = $"Data Source={ucChangePrix.selectedConnex};Initial Catalog={ucChangePrix.selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";
            CreateGridColumnsFromTable(gridControl1, connectionString);
        }
       
        public void CreateGridColumnsFromTable(GridControl gridControl, string connectionString)
        {
            // Connexion à la base de données
            string query = "SELECT CT_Intitule FROM P_CATTARIF WHERE CT_Intitule <> '' AND CT_Intitule IS NOT NULL";
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            using (SqlCommand command = new SqlCommand(query, connection))
            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
            {
                connection.Open();
                adapter.Fill(dataTable);
            }

            // Récupération de la vue (GridView)
            GridView gridView = gridControl.MainView as GridView;

            if (gridView != null)
            {
                gridView.Columns.Clear(); // Nettoyer les colonnes existantes

                foreach (DataRow row in dataTable.Rows)
                {
                    string columnName = row["CT_Intitule"].ToString();

                    // Ajouter une colonne pour chaque CT_Intitule
                    var gridColumn = gridView.Columns.AddVisible(columnName);
                    gridColumn.Caption = columnName;
                    gridColumn.FieldName = columnName; // Le FieldName doit correspondre à la source de données
                }
            }
        }
    }
}