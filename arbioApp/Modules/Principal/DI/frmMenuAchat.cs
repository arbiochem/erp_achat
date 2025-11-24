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
using arbioApp.Modules.Principal.DI._2_Documents;
using System.Data.SqlClient;
using BindingSource = System.Windows.Forms.BindingSource;

namespace arbioApp.Modules.Principal.DI
{
    public partial class frmMenuAchat : DevExpress.XtraEditors.XtraForm
    {
        public static string connectionString;
        private string dbPrincipale = ucDocuments.dbNamePrincipale;
        private string serveripPrincipale = ucDocuments.serverIpPrincipale;
        public frmMenuAchat()
        {
            InitializeComponent();
            connectionString = $"Server={serveripPrincipale};" +
                               $"Database={dbPrincipale};User ID=Dev;Password=1234;" +
                               $"TrustServerCertificate=True;Connection Timeout=120;";
        }

        public static int dotype = 10;
        public static int statut = 2; 
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if(rbAPA.Checked)
            {
                ucDocuments monUc = new ucDocuments();
                statut = 2; // accepté par défaut
                frmEditDocument editForm = new frmEditDocument(GetNextInvoiceNumber("APA"), monUc, 0);
                editForm.ShowDialog();
                dotype = 10;
                
            }
            
        }
        
        public static string GetNextInvoiceNumber(string prefix)
        {
            int year = DateTime.Now.Year;
            string query = "SELECT CurrentNumber FROM ARB_ACHAT_DOPIECE WHERE Prefix = @Prefix AND Year = @Year";
            
            // Récupérer le dernier numéro
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Prefix", prefix);
                    command.Parameters.AddWithValue("@Year", year);

                    object result = command.ExecuteScalar();
                    int currentNumber = result != DBNull.Value ? (int)result : 0;  // Si pas de résultat, commencer à 0

                    return $"{prefix}{year}{currentNumber:D4}";  
                }
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}