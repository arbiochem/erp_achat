using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using DevExpress.Utils.DirectXPaint;
using DevExpress.Utils.Extensions;
using DevExpress.XtraCharts.Designer.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BindingSource = System.Windows.Forms.BindingSource;

namespace arbioApp.Modules.Principal.DI
{
    public partial class frmMenuAchat : DevExpress.XtraEditors.XtraForm
    {
        private SqlDataAdapter adapter;
        private DataTable dt;
        private SqlCommandBuilder builder;
        public static string connectionString;
        private string dbPrincipale = ucDocuments.dbNamePrincipale;
        private string serveripPrincipale = ucDocuments.serverIpPrincipale;
        private bool DO_CREATE=false;
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
            if (rbAPA.Checked)
            {
                // Vérifie si l'utilisateur a le droit "CREATE" pour le type "APA"
                bool autorise = verifier_droit(rbAPA.Text, "CREATE");

                if (autorise)
                {
                    ucDocuments monUc = new ucDocuments();
                    statut = 2; // accepté par défaut
                    frmEditDocument editForm = new frmEditDocument(GetNextInvoiceNumber("APA"), monUc, 0);
                    editForm.lkDepot.EditValue = radioGroup1.EditValue;
                    editForm.ShowDialog();
                    dotype = 10;
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de créer un projet d'achat !",
                        "Création bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (rbABC.Checked){
                // Vérifie si l'utilisateur a le droit "CREATE" pour le type "APA"
                bool autorise = verifier_droit(rbABC.Text, "CREATE");

                if (autorise)
                {

                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de créer un bon de commande !",
                        "Création bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (rbABL.Checked)
            {
                // Vérifie si l'utilisateur a le droit "CREATE" pour le type "APA"
                bool autorise = verifier_droit(rbABL.Text, "CREATE");

                if (autorise)
                {

                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de créer un bon de livraison !",
                        "Création bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (rbAFC.Checked)
            {
                // Vérifie si l'utilisateur a le droit "CREATE" pour le type "APA"
                bool autorise = verifier_droit(rbAFC.Text, "CREATE");

                if (autorise)
                {

                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de créer une facture !",
                        "Création bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
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

        public static bool verifier_droit(string type, string cond)
        {
            bool verif = false;

            string connectionString = $"Server=SRV-ARB;" +
                                      $"Database=arbapp;User ID=Dev;Password=1234;" +
                                      $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "SELECT * FROM dbo.AUTORISATIONS_ACHAT WHERE mailUser = @mailUser AND type_document = @type_document";

            DataTable dt = new DataTable();

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@mailUser", FrmMdiParent.Username);
                        cmd.Parameters.AddWithValue("@type_document", type);

                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        SqlCommandBuilder builder = new SqlCommandBuilder(adapter);

                        adapter.Fill(dt);
                    }

                    foreach (DataRow row in dt.Rows)
                    {
                        // Ici, on suppose que cond contient le nom de la colonne à vérifier, par exemple "CREATE"
                        if (row.Table.Columns.Contains(cond) && row[cond] != DBNull.Value)
                        {
                            bool droit = Convert.ToBoolean(row[cond]);
                            if (droit)
                            {
                                verif = true;
                                break; 
                            }
                        } 
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors du remplissage : " + ex.Message);
            }

            return verif;
        }

        private void frmMenuAchat_Load(object sender, EventArgs e)
        {
            try
            {
                using (AppDbContext context = new AppDbContext())
                {
                    var depots = context.F_DEPOT.ToList();

                    radioGroup1.Properties.Items.Clear();

                    foreach (var depot in depots)
                    {
                        radioGroup1.Properties.Items.Add(
                            new RadioGroupItem(depot.DE_No, depot.DE_Intitule)   // Value = ID, Label = Nom
                        );
                    }

                    // affichage horizontal
                    radioGroup1.Properties.Columns = depots.Count;

                    // valeur par défaut = premier dépôt
                    if (depots.Any())
                        radioGroup1.EditValue = depots.First().DE_No;

                }
            }
            catch (System.Data.Entity.Infrastructure.DbUpdateException ex)
            {
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur");
            }

           
        }

        private void radioGroup1_EditValueChanged(object sender, EventArgs e)
        {

        }
    }
}