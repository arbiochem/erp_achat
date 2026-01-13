using DevExpress.XtraCharts;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frm_Etat_general : Form
    {
        private static string dbPrincipale = ucDocuments.dbNamePrincipale;
        private static string serveripPrincipale = ucDocuments.serverIpPrincipale;
        public frm_Etat_general()
        {
            InitializeComponent();
        }

        private void frm_Etat_general_Load(object sender, EventArgs e)
        {

        }

        private void load_fournisseur()
        {
            string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";

            cmbFournisseur.Properties.Items.Clear();
            cmbFournisseur.Properties.Items.Add("Tous");
            using (SqlConnection conn = new SqlConnection(connectionString2))
            {
                string query = @"
                SELECT DISTINCT CT_Intitule
                FROM F_COMPTET
                ORDER BY CT_Intitule";
                conn.Open();

                using (SqlCommand cmd = new SqlCommand(query, conn))
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        cmbFournisseur.Properties.Items.Add(dr.GetString(0));
                    }
                }
            }

            cmbFournisseur.SelectedIndex = -1;

        }

        private void load_etat_general()
        {
            string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";


            string query = "";

            if (string.IsNullOrEmpty(cmbFournisseur.Text) || cmbFournisseur.Text == "Tous")
            {
                query = @"
                SELECT *
                FROM Etat_global_achat
                WHERE [Date de commande] BETWEEN @dateDebut AND @dateFin
                ORDER BY [Date de LIVRAISON] DESC";
            }
            else
            {
                query = @"
                SELECT *
                FROM Etat_global_achat
                WHERE [Date de commande] BETWEEN @dateDebut AND @dateFin
                AND Fournisseur=@Fournisseur
                ORDER BY [Date de LIVRAISON] DESC";
            }

            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString2))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.Add("@dateDebut", SqlDbType.DateTime).Value = dtdate1.Value.Date; // minuit
                    cmd.Parameters.Add("@dateFin", SqlDbType.DateTime).Value = dtdate2.Value.Date.AddDays(1).AddTicks(-1); // 23:59:59.9999999
                    cmd.Parameters.Add("@Fournisseur", SqlDbType.VarChar).Value = cmbFournisseur.Text;

                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dt);
                    }
                }

                // 🔹 Binding au DataGrid / GridControl
                gdView.DataSource = dt;
                // ou DevExpress :
                GridView view = gdView.MainView as GridView;

                // Activer l'affichage du footer
                view.OptionsView.ShowFooter = true;

                view.Appearance.FooterPanel.BackColor = Color.Yellow;
                view.Appearance.FooterPanel.Font = new Font(view.Appearance.FooterPanel.Font, FontStyle.Bold);
                view.Appearance.FooterPanel.ForeColor = Color.DarkBlue;
                view.Appearance.FooterPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                for (int i = 3; i < view.Columns.Count; i++)
                {
                    GridColumn col = view.Columns[i];

                    // Centrer texte et en-tête
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                    // Format numérique avec 2 décimales
                    col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                    col.DisplayFormat.FormatString = "N2";

                    // Résumé pour "Prix de revient"
                    if (col.FieldName == "Prix de revient" || col.FieldName == "Montant total")
                    {
                        col.Summary.Clear();
                        col.Summary.Add(DevExpress.Data.SummaryItemType.Sum, col.FieldName, "{0:N2}");
                    }
                }

                if (dt.Rows.Count > 0)
                {
                    btnPrint.Enabled = true;
                }
                else
                {
                    btnPrint.Enabled = false;
                }

            }

        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (dtdate1.Value > dtdate2.Value)
            {
                MessageBox.Show("La plage de date est incorrecte!!!","Message d'erreur",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                load_etat_general();
            }
        }

        private void dtdate2_ValueChanged(object sender, EventArgs e)
        {
            load_fournisseur();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // Récupérer la GridView principale
            DevExpress.XtraGrid.Views.Grid.GridView view = gdView.MainView as DevExpress.XtraGrid.Views.Grid.GridView;

            if (view != null)
            {
                // Sauvegarder les largeurs actuelles
                int[] originalWidths = new int[view.Columns.Count];
                for (int i = 0; i < view.Columns.Count; i++)
                    originalWidths[i] = view.Columns[i].Width;

                // Ajuster les colonnes pour l'aperçu
                view.BestFitColumns();

                // Afficher l'aperçu avant impression
                gdView.ShowRibbonPrintPreview();

                // Restaurer les largeurs originales après fermeture de l'aperçu
                for (int i = 0; i < view.Columns.Count; i++)
                    view.Columns[i].Width = originalWidths[i];
            }

        }
    }
}
