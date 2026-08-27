using DevExpress.Map.Native;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Columns;
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
    public partial class frm_correspondance_articles : Form
    {
        private string dbPrincipale = ucDocuments.dbNamePrincipale;
        private string serveripPrincipale = ucDocuments.serverIpPrincipale;
        public frm_correspondance_articles()
        {
            InitializeComponent();
        }

        private void frm_correspondance_articles_Load(object sender, EventArgs e)
        {
            lister_article_sage();
            load_corres();
        }

        private void load_corres()
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            try
            {
                string query = "SELECT art.id,sage.AR_Design,art.article_achat FROM F_Corres_article as art INNER JOIN F_ARTICLE AS sage ON (sage.AR_Ref=art.reference_sage) ORDER BY art.article_achat asc";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // ✅ Correction : adapter avec paramètre
                    SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                   
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    gdCorres.DataSource = dt;

                    GridView view = gdCorres.MainView as GridView;

                    view.Columns["id"].Visible=false;
                    view.Columns["article_achat"].Caption = "Article dans Achat";
                    view.Columns["AR_Design"].Caption = "Article dans Sage";
                   
                    view.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    view.Appearance.HeaderPanel.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                    view.Appearance.HeaderPanel.ForeColor = Color.Blue;
                    view.OptionsBehavior.Editable = true;

                    // ✅ Supprimer le bouton s'il existe déjà
                    if (view.Columns["colAction"] != null)
                        view.Columns.Remove(view.Columns["colAction"]);

                    // ✅ Supprimer aussi l'ancien RepositoryItem associé, sinon il reste en mémoire
                    var oldRepoItem = gdCorres.RepositoryItems
                        .OfType<RepositoryItemButtonEdit>()
                        .FirstOrDefault(r => r.Name == "riColAction");
                    if (oldRepoItem != null)
                        gdCorres.RepositoryItems.Remove(oldRepoItem);

                    GridColumn colAction = new GridColumn();
                    colAction.Caption = "Actions";
                    colAction.VisibleIndex = 3;
                    colAction.UnboundType = DevExpress.Data.UnboundColumnType.String;
                    colAction.FieldName = "colAction";
                    view.Columns.Add(colAction);

                    RepositoryItemButtonEdit btnAction = new RepositoryItemButtonEdit();
                    btnAction.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
                    btnAction.Buttons.Clear();

                    btnAction.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton()
                    {
                        Caption = "✏️",
                        Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph,
                        Width = 40
                    });

                    // Bouton Supprimer
                    btnAction.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton()
                    {
                        Caption = "🗑",
                        Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph,
                        Width = 40
                    });

                    btnAction.ButtonClick += BtnAction_ButtonClick;

                    gdCorres.RepositoryItems.Add(btnAction); // ⚠️ souvent oublié
                    colAction.ColumnEdit = btnAction;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAction_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            GridView view = gdCorres.MainView as GridView; // adapte au nom de ton contrôle
            int rowHandle = view.FocusedRowHandle;

            // Index 0 = Modifier, Index 1 = Supprimer (ordre d'ajout ci-dessus)
            int buttonIndex = e.Button.Index;

            string articleAchatActuel = view.GetRowCellValue(rowHandle, "article_achat")?.ToString();
            string designSageActuel = view.GetRowCellValue(rowHandle, "AR_Design")?.ToString();

            if (buttonIndex == 0)
            {
                object id = view.GetRowCellValue(rowHandle, "id");
                OuvrirFormulaireEdition(articleAchatActuel, designSageActuel, id);
            }
            else if (buttonIndex == 1)
            {
                var result = MessageBox.Show("Confirmer la suppression ?", "Confirmation",
                                              MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    object id = view.GetRowCellValue(rowHandle, "id");
                    SupprimerLigne(id);
                }
            }

            load_corres();
        }

        private void OuvrirFormulaireEdition(string ancienArticleAchat, string ancienDesignSage, object id)
        {
            // Récupérer l'AR_Ref actuel correspondant au design Sage actuel
            string ancienRef = recupereridarticle(ancienDesignSage);

            using (Form editForm = new Form())
            {
                editForm.Text = "Modifier la correspondance";
                editForm.Size = new Size(480, 260);
                editForm.StartPosition = FormStartPosition.CenterParent;
                editForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                editForm.MaximizeBox = false;
                editForm.MinimizeBox = false;
                editForm.Padding = new Padding(20);

                int marge = 20;
                int largeurChamp = 420;

                Label lblArticleAchat = new Label()
                {
                    Text = "Article dans Achat :",
                    Location = new Point(marge, marge),
                    AutoSize = true
                };
                TextBox txtArticleAchatEdit = new TextBox()
                {
                    Text = ancienArticleAchat,
                    Location = new Point(marge, marge + 25),
                    Width = largeurChamp
                };

                Label lblArticleSage = new Label()
                {
                    Text = "Article dans Sage :",
                    Location = new Point(marge, marge + 65),
                    AutoSize = true
                };
                ComboBox cmbArticleSageEdit = new ComboBox()
                {
                    Location = new Point(marge, marge + 90),
                    Width = largeurChamp,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };

                // Charger la liste des articles Sage
                string connectionString = $"Server={serveripPrincipale};" +
                                         $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                         $"TrustServerCertificate=True;Connection Timeout=120;";

                try
                {
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string q = "SELECT DISTINCT AR_Design FROM F_ARTICLE WHERE AR_Sommeil=0 ORDER BY AR_Design asc";
                        using (SqlCommand cmd = new SqlCommand(q, conn))
                        {
                            SqlDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                cmbArticleSageEdit.Items.Add(reader["AR_Design"].ToString());
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur de chargement des articles Sage : {ex.Message}", "Erreur",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                cmbArticleSageEdit.SelectedItem = ancienDesignSage;

                Button btnEnregistrer = new Button()
                {
                    Text = "Enregistrer",
                    Location = new Point(marge + 90, marge + 140),
                    Width = 110,
                    Height = 32
                };
                Button btnAnnuler = new Button()
                {
                    Text = "Annuler",
                    Location = new Point(marge + 210, marge + 140),
                    Width = 110,
                    Height = 32
                };

                btnEnregistrer.Click += (s, ev) =>
                {
                    if (string.IsNullOrWhiteSpace(txtArticleAchatEdit.Text) || cmbArticleSageEdit.SelectedItem == null)
                    {
                        MessageBox.Show("Veuillez renseigner tous les champs.", "Erreur",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    string nouveauDesign = cmbArticleSageEdit.SelectedItem.ToString();
                    string nouvelleRef = recupereridarticle(nouveauDesign);

                    ModifierCorrespondance(txtArticleAchatEdit.Text.Trim(), nouvelleRef, id);

                    editForm.DialogResult = DialogResult.OK;
                    editForm.Close();
                };

                btnAnnuler.Click += (s, ev) =>
                {
                    editForm.DialogResult = DialogResult.Cancel;
                    editForm.Close();
                };

                editForm.Controls.Add(lblArticleAchat);
                editForm.Controls.Add(txtArticleAchatEdit);
                editForm.Controls.Add(lblArticleSage);
                editForm.Controls.Add(cmbArticleSageEdit);
                editForm.Controls.Add(btnEnregistrer);
                editForm.Controls.Add(btnAnnuler);

                editForm.ShowDialog(this);
            }
        }

        private void ModifierCorrespondance(string nouvelArticleAchat, string nouvelleRef, object id)
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "UPDATE F_Corres_article SET article_achat = @nouvelArticleAchat, reference_sage = @nouvelleRef " +
                           "WHERE id = @id";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.Add("@id", SqlDbType.Int, 50).Value = id;
                        cmd.Parameters.Add("@nouvelArticleAchat", SqlDbType.VarChar, 50).Value = nouvelArticleAchat;
                        cmd.Parameters.Add("@nouvelleRef", SqlDbType.VarChar, 50).Value = nouvelleRef;

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Modification réussie.", "Succès",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);

                load_corres();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la modification : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SupprimerLigne(object cond)
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "DELETE FROM F_Corres_article WHERE id=@id";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.Add("@id", SqlDbType.Int, 50).Value = int.Parse(cond.ToString());
                        
                            cmd.ExecuteNonQuery();

                            MessageBox.Show("Suppression réussie.", "Succès",
                               MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (SqlException ex)
            {
                // Idéalement, logger ex.ToString() (avec stack trace) dans un fichier/log
                MessageBox.Show("Une erreur est survenue lors de l'enregistrement. Veuillez réessayer.",
                                "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur inattendue : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void lister_article_sage()
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "SELECT DISTINCT AR_Design FROM F_ARTICLE WHERE AR_Sommeil=0 ORDER BY AR_Design asc";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataReader reader = cmd.ExecuteReader();

                        cmbartSAGE.Properties.Items.Clear();

                        while (reader.Read())
                        {
                            cmbartSAGE.Properties.Items.Add(reader["AR_Design"].ToString());
                        }
                    }
                }

                cmbartSAGE.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur de connexion : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbartSAGE_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbl_code_article.Text = "";
            string recup=cmbartSAGE.SelectedItem.ToString();
            lbl_code_article.Text = recupereridarticle(recup);
        }

        private string recupereridarticle(string cond)
        {
            string val = "";

            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "SELECT DISTINCT AR_Ref FROM F_ARTICLE WHERE AR_Design LIKE '%' + @cond + '%'";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@cond", cond);
                        SqlDataReader reader = cmd.ExecuteReader();

                        val="";

                        while (reader.Read())
                        {
                            val=reader["AR_Ref"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur de connexion : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return val;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            // Validation avant tout appel base de données
            if (string.IsNullOrWhiteSpace(txtartachat.Text) || string.IsNullOrWhiteSpace(lbl_code_article.Text))
            {
                MessageBox.Show("Veuillez renseigner tous les champs.", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string query = "INSERT INTO F_Corres_article (article_achat, reference_sage) " +
                           "VALUES (@article_achat, @reference_sage)";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.Add("@article_achat", SqlDbType.VarChar, 50).Value = txtartachat.Text.Trim();
                        cmd.Parameters.Add("@reference_sage", SqlDbType.VarChar, 50).Value = lbl_code_article.Text.Trim();

                        if (verifier_existence(lbl_code_article.Text.Trim()))
                        {
                            MessageBox.Show("Cet article SAGE possède déjà un correspondant dans Achat", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            cmd.ExecuteNonQuery();

                            MessageBox.Show("Enregistrement réussi.", "Succès",
                               MessageBoxButtons.OK, MessageBoxIcon.Information);

                            load_corres();
                        }

                    }
                }
            }
            catch (SqlException ex)
            {
                // Idéalement, logger ex.ToString() (avec stack trace) dans un fichier/log
                MessageBox.Show("Une erreur est survenue lors de l'enregistrement. Veuillez réessayer.",
                                "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur inattendue : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool verifier_existence(String cond)
        {
            bool b_test = false;

            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "SELECT DISTINCT reference_sage FROM F_Corres_article WHERE reference_sage =@reference";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@reference", cond);
                        SqlDataReader reader = cmd.ExecuteReader();

                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0) && !string.IsNullOrWhiteSpace(reader.GetString(0)))
                            {
                                b_test = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur de connexion : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return b_test;
        }
    }
}
