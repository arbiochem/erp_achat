using DevExpress.DataProcessing;
using DevExpress.Xpo;
using DevExpress.Xpo.DB.Helpers;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Exception = System.Exception;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frm_correspondance : Form
    {
        private string dbPrincipale = ucDocuments.dbNamePrincipale;
        private string serveripPrincipale = ucDocuments.serverIpPrincipale;

        public frm_correspondance()
        {
            InitializeComponent();
        }

        private void frm_correspondance_Load(object sender, EventArgs e)
        {
            load_fournisseur();
            load_correspondance();
        }


        private void load_correspondance()
        {
            gd_data.DataSource = null;
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            try
            {
                string query = "SELECT DISTINCT num_piece, fourn, fourns FROM F_Corres WHERE num_piece = @num_piece ORDER BY fourn ASC";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // ✅ Correction : adapter avec paramètre
                    SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                    adapter.SelectCommand.Parameters.AddWithValue("@num_piece", txtnumpiece.Text.Trim());

                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    gd_data.DataSource = dt;

                    GridView view = gd_data.MainView as GridView;

                    view.Columns["num_piece"].Caption = "Numéro de pièce";
                    view.Columns["num_piece"].VisibleIndex = 0;
                    view.Columns["fourn"].Caption = "Fournisseur 1";
                    view.Columns["fourn"].VisibleIndex = 1;
                    view.Columns["fourns"].Caption = "Fournisseur 2";
                    view.Columns["fourns"].VisibleIndex = 2;

                    view.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    view.Appearance.HeaderPanel.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                    view.Appearance.HeaderPanel.ForeColor = Color.Blue;
                    view.OptionsBehavior.Editable = false;

                    // ✅ Supprimer le bouton s'il existe déjà
                    if (view.Columns["colSupprimer"] != null)
                        view.Columns.Remove(view.Columns["colSupprimer"]);

                    // ✅ Bouton en dernière position (index 3)
                    GridColumn btnSupprimer = new GridColumn();
                    btnSupprimer.Caption = "Action";
                    btnSupprimer.VisibleIndex = 3;
                    btnSupprimer.UnboundType = DevExpress.Data.UnboundColumnType.String;
                    btnSupprimer.FieldName = "colSupprimer";
                    view.Columns.Add(btnSupprimer);

                    RepositoryItemButtonEdit btnEdit = new RepositoryItemButtonEdit();
                    btnEdit.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
                    btnEdit.Buttons.Clear();
                    btnEdit.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton()
                    {
                        Caption = "🗑 Supprimer",
                        Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph,
                        Width = 80
                    });

                    gd_data.MouseClick -= Gd_data_MouseClick;
                    gd_data.MouseClick += Gd_data_MouseClick;

                    gd_data.RepositoryItems.Add(btnEdit);
                    btnSupprimer.ColumnEdit = btnEdit;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Gd_data_MouseClick(object sender, MouseEventArgs e)
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            GridView view = gd_data.MainView as GridView;

            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            if (hitInfo.InRowCell && hitInfo.Column?.FieldName == "colSupprimer")
            {
                int rowHandle = hitInfo.RowHandle;
                if (rowHandle < 0) return;

                string num_piece = view.GetRowCellValue(rowHandle, "num_piece")?.ToString();
                string fourn = view.GetRowCellValue(rowHandle, "fourn")?.ToString();
                string fourns = view.GetRowCellValue(rowHandle, "fourns")?.ToString();

                var result = MessageBox.Show(
                    $"Voulez-vous supprimer :\n{fourn} - {fourns} pour le dossier {num_piece} ?",
                    "Confirmation",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        using (SqlConnection connDel = new SqlConnection(connectionString))
                        {
                            connDel.Open();
                            string deleteQuery = "DELETE FROM F_Corres WHERE fourn = @fourn AND fourns = @fourns AND num_piece = @num_piece";
                            using (SqlCommand cmdDel = new SqlCommand(deleteQuery, connDel))
                            {
                                cmdDel.Parameters.AddWithValue("@num_piece", num_piece);
                                cmdDel.Parameters.AddWithValue("@fourn", fourn);
                                cmdDel.Parameters.AddWithValue("@fourns", fourns);
                                cmdDel.ExecuteNonQuery();
                            }
                        }

                        view.DeleteRow(rowHandle);

                        MessageBox.Show("Supprimé avec succès !", "Succès",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erreur suppression : {ex.Message}", "Erreur",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void load_fournisseur()
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM_ACHAT;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "SELECT DISTINCT CT_INTITULE FROM F_COMPTET WHERE ISNULL(CT_INTITULE, '') <> '' AND CT_Sommeil=0 and CT_Type = 1 ORDER BY CT_INTITULE asc";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataReader reader = cmd.ExecuteReader();

                        cmb_fourn.Properties.Items.Clear();
                        cmb_fourns.Properties.Items.Clear();

                        while (reader.Read())
                        {
                            cmb_fourn.Properties.Items.Add(reader["CT_INTITULE"].ToString());
                            cmb_fourns.Properties.Items.Add(reader["CT_INTITULE"].ToString());
                        }
                    }
                }

                cmb_fourn.Refresh();
                cmb_fourns.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur de connexion : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        Boolean tester_correspondance()
        {
            bool d_val = false;

            string connectionString = $"Server={serveripPrincipale};" +
                                     $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                     $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "SELECT * FROM F_Corres WHERE fourn=@fourn AND fourns=@fourns AND num_piece=@num_piece";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@fourn", cmb_fourn.Text.ToString());
                        cmd.Parameters.AddWithValue("@fourns", cmb_fourns.Text.ToString());
                        cmd.Parameters.AddWithValue("@num_piece", txtnumpiece.Text.ToString());

                        SqlDataReader reader = cmd.ExecuteReader();

                        if (reader.HasRows) 
                        {
                            d_val = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur de connexion : {ex.Message}", "Erreur",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return d_val;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (cmb_fourn.Text == "" || cmb_fourns.Text == "")
            {
                MessageBox.Show("Données vides, veuillez remplir les 2 champs","Message d'erreur",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (cmb_fourn.Text == cmb_fourns.Text)
                {
                    MessageBox.Show("Les 2 fournisseurs ne peuvent pas être identiques", "Message d'erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (tester_correspondance())
                    {
                        MessageBox.Show("Cette correspondace existe déjà dans la base", "Message d'erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        string connectionString = $"Server={serveripPrincipale};" +
                                             $"Database=ARBIOCHEM;User ID=Dev;Password=1234;" +
                                             $"TrustServerCertificate=True;Connection Timeout=120;";

                        string query = "INSERT INTO F_Corres(fourn,fourns,num_piece) VALUES(@fourn,@fourns,@num_piece)";

                        try
                        {
                            using (SqlConnection connection = new SqlConnection(connectionString))
                            {
                                connection.Open();
                                using (SqlCommand insertCmd = new SqlCommand(query, connection))
                                {
                                    insertCmd.Parameters.AddWithValue("@fourn", cmb_fourn.Text.ToString());
                                    insertCmd.Parameters.AddWithValue("@fourns", cmb_fourns.Text.ToString());
                                    insertCmd.Parameters.AddWithValue("@num_piece", txtnumpiece.Text.ToString());

                                    insertCmd.ExecuteNonQuery();
                                }
                                connection.Close();

                                MessageBox.Show("Données insérées avec succès!!!", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                cmb_fourn.SelectedIndex = -1;
                                cmb_fourns.SelectedIndex = -1;

                                load_correspondance();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Erreur de connexion : {ex.Message}", "Erreur",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }
    }
}
