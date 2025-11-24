using arbioApp.Modules.Principal.GESTION_COMMERCIALE;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace arbioApp.Modules.Principal.DSI
{
    public partial class FrmNewUser : DevExpress.XtraEditors.XtraForm
    {
        public FrmNewUser()
        {
            InitializeComponent();
            server = ucReleaseUser.selectedConnex;
            catalog = ucReleaseUser.selectedSite;
            getProtGuid();
        }

        private string server;
        private string catalog;
        private void getProtGuid()
        {
            string connectionString = $"Data Source={server};Initial Catalog={catalog};User ID=Dev;Password=1234;TrustServerCertificate=True";
            string query = @"SELECT DISTINCT 
                              dbo.F_PROTECTIONCIAL.PROT_User,
                              dbo.F_PROTECTIONCIAL.PROT_Guid,
                              dbo.F_DEPOT_DEDIE.ID
                            FROM
                              dbo.F_PROTECTIONCIAL
                              LEFT OUTER JOIN dbo.F_DEPOT_DEDIE ON (dbo.F_PROTECTIONCIAL.PROT_Guid = dbo.F_DEPOT_DEDIE.PROT_Guid)    
                            WHERE 
                              dbo.F_DEPOT_DEDIE.ID IS NULL";
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            using (SqlCommand command = new SqlCommand(query, connection))
            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
            {
                connection.Open();
                adapter.Fill(dataTable);
            }

            lookUpEdit1.Properties.DataSource = dataTable;
            lookUpEdit1.Properties.DisplayMember = "PROT_User"; 
            lookUpEdit1.Properties.ValueMember = "PROT_Guid"; 
            lookUpEdit1.Properties.NullText = "Rattachez un utilisateur";
            //InsertNewUser(connectionString,lookUpEdit1.EditValue.ToString() );
        }

        private void InsertNewUser(string cs, string protguid)
        {
            
            string query = @"INSERT INTO F_DEPOT_DEDIE (DE_NO, PROT_Guid, AUTHORIZED)
                                    SELECT
                                        d.DE_NO,
                                        p.PROT_Guid,
                                        0 -- AUTHORIZED par défaut
                                    FROM
                                        F_DEPOT d
                                    CROSS JOIN
                                        F_PROTECTIONCIAL p
                                    WHERE
                                        NOT EXISTS (
                                            SELECT 1
                                            FROM F_DEPOT_DEDIE dd
                                            WHERE dd.DE_NO = d.DE_NO
                                              AND dd.PROT_Guid = p.PROT_Guid
                                        )
                                    AND
                                    p.PROT_Guid = @PROT_Guid";

            using (SqlConnection connection = new SqlConnection(cs))
            {
                try
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@PROT_Guid", protguid);
                        int rowsAffected = command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            string connectionString = $"Data Source={server};Initial Catalog={catalog};User ID=Dev;Password=1234;TrustServerCertificate=True";

            InsertNewUser(connectionString,lookUpEdit1.EditValue.ToString() );
        }

    }
    
}