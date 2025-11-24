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

namespace arbioApp.Modules.Principal.Theme
{
    public partial class FrmTheme : DevExpress.XtraEditors.XtraForm
    {
        public string Username { get; private set; }
        public string UserRole { get; private set; }

        public bool success;
        public string leserveur { get; set; }
        public string labase { get; set; }
        public FrmTheme()
        {
            InitializeComponent();
            UserRole = FrmMdiParent.UserRole;
            leserveur = FrmMdiParent.DataSourceNameValueParent;
            Username = FrmMdiParent.Username;
            LoadThemes();
        }
        private void LoadThemes()
        {
           
            using (SqlConnection connection = new SqlConnection(
                           $"Data Source={leserveur};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True"))
            {
                string query = "SELECT Id_Theme, Intitule_Theme, lOption FROM T_Theme";

                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                gridControl1.DataSource = dataTable;
            }
        }

        private void gridView1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName == "lOption")
            {
                GridView view = sender as GridView;
               
                // Obtenir la valeur de la cellule modifiée
                bool isChecked = Convert.ToBoolean(e.Value);

                if (isChecked) // Si la case est cochée
                {
                    // Parcourir toutes les lignes et décocher les autres
                    for (int i = 0; i < view.RowCount; i++)
                    {
                        if (i != e.RowHandle) // Sauf la ligne actuellement modifiée
                        {
                            view.SetRowCellValue(i, "LOption", false);
                        }
                    }
                    // Appliquer le thème correspondant
                    string selectedTheme = view.GetRowCellValue(e.RowHandle, "Intitule_Theme").ToString();
                    ApplyTheme(selectedTheme); 
                    //int themeId = Convert.ToInt32(view.GetRowCellValue(e.RowHandle, "id_theme"));
                    int themeId = e.RowHandle;
                    idtheme = themeId;
                    SaveSelectedTheme(themeId);
                }
            }
        }
        int idtheme;
        public static void ApplyTheme(string themeName)
        {
            try
            {
                // Changer le thème avec DevExpress
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(themeName);
                MessageBox.Show($"Thème '{themeName}' appliqué avec succès.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'application du thème : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FrmTheme_Load(object sender, EventArgs e)
        {

        }

        private void gridView1_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {
            GridView view = sender as GridView;
            if (e.Action == System.ComponentModel.CollectionChangeAction.Add && view.GetSelectedRows().Length > 1)
                view.ClearSelection();
            view.SelectRow(view.GetRowHandle(e.ControllerRow));
        }
        private void SaveSelectedTheme(int themeId)
        {
            using (SqlConnection connection = new SqlConnection(
                            $"Data Source={leserveur};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True"))
            { 
                string query = @"
            UPDATE T_Theme SET lOption = 0; -- Réinitialiser toutes les options
            UPDATE T_Theme SET lOption = 1 WHERE id_theme = @id_theme; -- Sélectionner le thème";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id_Theme", themeId);
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }

        private void hyperlinkSaveTheme_Click(object sender, EventArgs e)
        {
            SaveSelectedTheme(idtheme);
            this.Close();
        }
    }
}