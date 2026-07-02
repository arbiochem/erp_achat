using DevExpress.CodeParser;
using DevExpress.XtraCharts.Native;
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
    public partial class frmArticles : Form
    {
        private static string serveripPrincipale = ucDocuments.serverIpPrincipale;

        private static string connectionString = $"Server={serveripPrincipale};Database=TRANSIT;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";

        private string id_fournisseur;
        private string id_article;

        public frmArticles()
        {
            InitializeComponent();
            cmbArticle.EditValueChanged += cmbArticle_EditValueChanged;
        }

        private void btnEnregistrer_Click(object sender, EventArgs e)
        {
            try
            {
                string connectionString = $"Server={serveripPrincipale};Database=TRANSIT;" +
                                $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                $"Connection Timeout=240;";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    decimal existingQte = 0;

                    string existingVal = @"
                    SELECT COUNT(*)
                    FROM F_ARTFOURNISS
                    WHERE AR_Ref = @AR_Ref AND CT_Num = @CT_Num";

                    using (SqlCommand checkCmd = new SqlCommand(existingVal, connection))
                    {
                        checkCmd.Parameters.AddWithValue("@AR_Ref", (object)id_article ?? DBNull.Value);
                        checkCmd.Parameters.AddWithValue("@CT_Num", (object)id_fournisseur ?? DBNull.Value);

                        var result = checkCmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                            existingQte = Convert.ToDecimal(result);
                    }

                    if (cmbArticle.SelectedIndex == -1)
                    {
                        // L'association existe déjà, ne rien insérer
                        MessageBox.Show("Aucun article sélectionné.",
                            "Message d'erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        if (existingQte > 0)
                        {
                            // L'association existe déjà, ne rien insérer
                            MessageBox.Show("Cet article est déjà associé à ce fournisseur.",
                                "Message d'erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            string insertSql = "INSERT INTO F_ARTFOURNISS (AR_Ref, CT_Num, cbCreationUser) VALUES (@AR_Ref, @CT_Num,@cbCreationUser)";
                            using (SqlCommand insertCmd = new SqlCommand(insertSql, connection))
                            {
                                insertCmd.Parameters.AddWithValue("@AR_Ref", (object)id_article ?? DBNull.Value);
                                insertCmd.Parameters.AddWithValue("@CT_Num", (object)id_fournisseur ?? DBNull.Value);
                                insertCmd.Parameters.AddWithValue("@cbCreationUser", "C6D19B13-D5FE-4A92-8F5C-165BEA801BD1");

                                insertCmd.ExecuteNonQuery();

                                MessageBox.Show("Cet article est bien associé à ce fournisseur.",
                                "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                this.Hide();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur");
            }
        }

        private void frmArticles_Load(object sender, EventArgs e)
        {
            load_article();
            String cond = this.Text.Replace("Ajout article pour ", "");
            id_fournisseur = recuperer_code_fournisseur(cond);
        }

        private void load_article()
        {

            cmbArticle.Properties.Items.Clear(); 

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string existarticle = @"
                SELECT DISTINCT AR_Design
                FROM F_ARTICLE
                ORDER BY AR_Design ASC";

                using (SqlCommand checkCmd = new SqlCommand(existarticle, connection))
                using (SqlDataReader reader = checkCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string design = reader["AR_Design"]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(design))
                            cmbArticle.Properties.Items.Add(design); 
                    }
                }
            }

            cmbArticle.Refresh();
        }

        private string recuperer_code_fournisseur(string cond)
        {
            string recup = null;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string existfournisseur = @"
                SELECT CT_Num
                FROM F_COMPTET
                WHERE CT_INTITULE LIKE @ctintitule";

                using (SqlCommand checkCmd = new SqlCommand(existfournisseur, connection))
                {
                    checkCmd.Parameters.AddWithValue("@ctintitule", "%" + cond.Trim() + "%");
                    var result = checkCmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                        recup = result.ToString().Trim();
                }

            }

            return recup;
        }

        private string recuperer_code_article(string cond)
        {
            string recup = null;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string existarticle = @"
                SELECT AR_Ref
                FROM F_ARTICLE
                WHERE AR_Design LIKE @ardesign";

                using (SqlCommand checkCmd = new SqlCommand(existarticle, connection))
                {
                    checkCmd.Parameters.AddWithValue("@ardesign", "%" + cond.Trim() + "%");
                    var result = checkCmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                        recup = result.ToString().Trim();
                }

            }

            return recup;
        }

        private void cmbArticle_EditValueChanged(object sender, EventArgs e)
        {
            string cond = cmbArticle.EditValue?.ToString();
            if (string.IsNullOrEmpty(cond)) return;

            id_article = recuperer_code_article(cond);
            cmbArticle.Refresh();
        }
    }
}
