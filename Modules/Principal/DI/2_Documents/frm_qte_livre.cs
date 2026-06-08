using arbioApp.Models;
using DevExpress.CodeParser;
using DevExpress.DataAccess.DataFederation;
using DevExpress.DataProcessing.InMemoryDataProcessor;
using DevExpress.DataProcessing.InMemoryDataProcessor.GraphGenerator;
using DevExpress.PivotGrid.OLAP;
using DevExpress.Utils.Text;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraReports.UI;
using DevExpress.XtraSpreadsheet.DocumentFormats.Xlsb;
using Microsoft.Reporting.Map.WebForms.BingMaps;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows.Forms;
using static DevExpress.Xpo.DB.DataStoreLongrunnersWatch;
using static System.Data.Entity.Infrastructure.Design.Executor;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frm_qte_livre : Form
    {
        private static string dbPrincipale = ucDocuments.dbNamePrincipale;
        private static string serveripPrincipale = ucDocuments.serverIpPrincipale;
        private static string connectionString = $"Server={serveripPrincipale};Database=TRANSIT;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";
        public int compter=0;

        public frm_qte_livre()
        {
            
            InitializeComponent();
            txtLot.Text = "";
            txtnumerolot.Text = "";
        }

        public static bool tester_dopiece(string cond)
        {
            bool verif = false;

            string query = "SELECT COUNT(1) FROM dbo.F_DOCENTETE WHERE Do_piece = @dopiece";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@dopiece", cond);
                        int count = (int)cmd.ExecuteScalar();
                        if(count > 0) verif=true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la vérification : " + ex.Message);
            }

            return verif;
        }

        private void btn_valider_Click(object sender, EventArgs e)
        {
            if (txtQteDepart.Text != txtQteLivrer.Text)
            {
                if (Convert.ToDecimal(txtQteLivrer.Text) > Convert.ToDecimal(txtQteDepart.Text))
                {
                    MessageBox.Show($"La quantité à livrer ne doit pas dépasser à la quantité demandée!!!!", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtQteLivrer.Text = "";
                    txtQteLivrer.Focus();
                }
                else
                {
                    string doPieceOriginal = this.Text.Replace("Traitement du document N° ", "").Trim();

                    // Tronquer à la taille max de la colonne (ex: 13)
                    int maxLen = 17; // adaptez selon votre colonne
                    if (doPieceOriginal.Length > maxLen)
                        doPieceOriginal = doPieceOriginal.Substring(0, maxLen);

                    // Parser la quantité correctement
                    if (!decimal.TryParse(
                            txtQteLivrer.Text.Replace(",", "."),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out decimal qte))
                    {
                        MessageBox.Show("Quantité invalide.", "Erreur",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    decimal qtt = Convert.ToDecimal(txtQteDepart.Text) - Convert.ToDecimal(txtQteLivrer.Text);

                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        if (Convert.ToDecimal(txtQteLivrer.Text) > 0)
                        {
                            string query = "";
                            if (!tester_dopiece(doPieceOriginal.Replace("AFA", "ABR")))
                            {
                                query = @"
                            UPDATE dbo.F_DOCLIGNE SET DL_Qte=@qte,DL_QteBC=@qte,DL_QteBL=@qte,EU_Qte=@qte,DL_QtePL=@qte,DL_QteDE=@qte,QteLivre=@qtelivre WHERE DO_Piece = @DoPiece AND AR_Ref=@arref";
                            }
                            else
                            {
                                query = @"
                                UPDATE dbo.F_DOCLIGNE SET 
                                    DL_Qte   = CAST(ISNULL(TRY_CAST(@qte AS DECIMAL(18,2)), 0) AS DECIMAL(18,2)),
                                    DL_QteBC = CAST(ISNULL(TRY_CAST(@qte AS DECIMAL(18,2)), 0) AS DECIMAL(18,2)),
                                    DL_QteBL = CAST(ISNULL(TRY_CAST(@qte AS DECIMAL(18,2)), 0) AS DECIMAL(18,2)),
                                    EU_Qte   = CAST(ISNULL(TRY_CAST(@qte AS DECIMAL(18,2)), 0) AS DECIMAL(18,2)),
                                    DL_QtePL = CAST(ISNULL(TRY_CAST(@qte AS DECIMAL(18,2)), 0) AS DECIMAL(18,2)),
                                    DL_QteDE = CAST(ISNULL(TRY_CAST(@qte AS DECIMAL(18,2)), 0) AS DECIMAL(18,2)),
                                    QteLivre += @qtelivre 
                                WHERE DO_Piece = @DoPiece AND AR_Ref = @arref";

                            }
                            conn.Open();
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@DoPiece", doPieceOriginal);
                                cmd.Parameters.AddWithValue("@arref", txtRef.Text.Trim());
                                cmd.Parameters.AddWithValue("@qte", qtt);
                                cmd.Parameters.AddWithValue("@qtelivre", Convert.ToDecimal(txtQteLivrer.Text));

                                new SqlCommand("DISABLE TRIGGER ALL ON F_DOCLIGNE", conn).ExecuteNonQuery();
                                cmd.ExecuteNonQuery();
                                new SqlCommand("ENABLE TRIGGER ALL ON F_DOCLIGNE", conn).ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            else
            {
                if (Convert.ToDecimal(txtQteLivrer.Text) > 0)
                {
                    string doPieceOriginal = this.Text.Replace("Traitement du document N° ", "").Trim();

                    AppDbContext context = new AppDbContext();
                    var test = context.F_DOCLIGNE.Where(p => p.DO_Piece.Contains(doPieceOriginal.Replace("AFA", "ABR"))).FirstOrDefault();

                    if (test == null)
                    {
                        var test1 = context.F_DOCLIGNE.Where(p => p.DO_Piece.Contains(doPieceOriginal)).FirstOrDefault();

                        if (test1 == null)
                        {
                            using (SqlConnection conn = new SqlConnection(connectionString))
                            {

                                conn.Open();
                                string querys = @"
                                DISABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                                UPDATE dbo.F_DOCLIGNE SET 
                                    DL_QteBC  = TRY_CAST(DL_Qte AS DECIMAL(18,2)) ,
                                    DL_QteBL  = TRY_CAST(DL_Qte AS DECIMAL(18,2)),
                                    DL_QtePL  = TRY_CAST(DL_Qte AS DECIMAL(18,2)),
                                    EU_Qte    = TRY_CAST(EU_Qte AS DECIMAL(18,2)),
                                    DL_QteDE  = TRY_CAST(DL_Qte AS DECIMAL(18,2)),
                                    QteLivre    = TRY_CAST(DL_Qte AS DECIMAL(18,2))
                                WHERE Do_Piece = @DocPiece AND AR_Ref = @arref;
                                ENABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                            ";

                                using (SqlCommand cmds = new SqlCommand(querys, conn))
                                {
                                    cmds.Parameters.AddWithValue("@DocPiece", doPieceOriginal);
                                    cmds.Parameters.AddWithValue("@arref", txtRef.Text);
                                    cmds.ExecuteNonQuery();
                                }
                            }
                        }
                        else
                        {
                            using (SqlConnection conn = new SqlConnection(connectionString))
                            {

                                conn.Open();
                                string querys = @"
                                DISABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                                UPDATE dbo.F_DOCLIGNE SET 
                                    DL_QteBC  = TRY_CAST(DL_Qte AS DECIMAL(18,2)) ,
                                    DL_QteBL  = TRY_CAST(DL_Qte AS DECIMAL(18,2)),
                                    DL_QtePL  = TRY_CAST(DL_Qte AS DECIMAL(18,2)),
                                    EU_Qte    = TRY_CAST(EU_Qte AS DECIMAL(18,2)),
                                    DL_QteDE  = TRY_CAST(DL_Qte AS DECIMAL(18,2)),
                                    QteLivre    = TRY_CAST(DL_Qte AS DECIMAL(18,2)),
                                    DO_Piece=DO_Piece+'_'
                                WHERE Do_Piece = @DocPiece AND AR_Ref = @arref;
                                ENABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                            ";

                                using (SqlCommand cmds = new SqlCommand(querys, conn))
                                {
                                    cmds.Parameters.AddWithValue("@DocPiece", doPieceOriginal);
                                    cmds.Parameters.AddWithValue("@arref", txtRef.Text);
                                    cmds.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    else
                    {
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {

                            conn.Open();
                            string querys = @"
                        DISABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                        UPDATE dbo.F_DOCLIGNE SET 
                            DL_Qte    = TRY_CAST(DL_Qte AS DECIMAL(18,2))     + TRY_CAST(QteLivre AS DECIMAL(18,2)),
                            DL_QteBC  = TRY_CAST(DL_QteBC AS DECIMAL(18,2))   + TRY_CAST(QteLivre AS DECIMAL(18,2)),
                            DL_QteBL  = TRY_CAST(DL_QteBL AS DECIMAL(18,2))   + TRY_CAST(QteLivre AS DECIMAL(18,2)),
                            DL_QtePL  = TRY_CAST(DL_QtePL AS DECIMAL(18,2))   + TRY_CAST(QteLivre AS DECIMAL(18,2)),
                            EU_Qte    = TRY_CAST(EU_Qte AS DECIMAL(18,2))     + TRY_CAST(QteLivre AS DECIMAL(18,2)),
                            DL_QteDE  = TRY_CAST(DL_QteDE AS DECIMAL(18,2))   + TRY_CAST(QteLivre AS DECIMAL(18,2)),
                            QteLivre    = TRY_CAST(DL_Qte AS DECIMAL(18,2))    + TRY_CAST(QteLivre AS DECIMAL(18,2)),
                            DO_Piece=DO_Piece+'_'
                        WHERE Do_Piece = @DocPiece AND AR_Ref = @arref;
                        ENABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                    ";

                            using (SqlCommand cmds = new SqlCommand(querys, conn))
                            {
                                cmds.Parameters.AddWithValue("@DocPiece", doPieceOriginal);
                                cmds.Parameters.AddWithValue("@arref", txtRef.Text);
                                cmds.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }

            this.Hide();
            /*try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = con;

                        if (!decimal.TryParse(
                            txtQteLivrer.Text.Trim(),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out decimal qte))
                        {
                            MessageBox.Show("Quantité invalide");
                            return;
                        }

                        if (!string.IsNullOrWhiteSpace(txtLot.Text))
                        {
                            cmd.Parameters.Clear();

                            cmd.CommandText = @"
                                    ALTER TABLE F_LotSerie NOCHECK CONSTRAINT ALL;";
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = @"
                                    INSERT INTO F_LotSerie
                                    (AR_Ref, LS_NoSerie, LS_Qte, LS_QteRestant, LS_Peremption, DE_No,LS_Fabrication,cbCreationUser,DL_NoIn,LS_LotEpuise,LS_MvtStock,DL_NoOut,LS_Complement,LS_QteRes)
                                    VALUES
                                    (@AR_Ref, @LS_NoSerie, @Qte, @Qte, @Peremption, @DE_No,@lsfabrication, @user,@dlnoin,0,1,0,@lscomplement,0)";

                            cmd.Parameters.AddWithValue("@AR_Ref", txtRef.Text);
                            cmd.Parameters.AddWithValue("@LS_NoSerie", txtLot.Text);
                            cmd.Parameters.Add("@Qte", qte);
                            cmd.Parameters.AddWithValue("@Peremption", Convert.ToDateTime(dtperemption.Text));
                            cmd.Parameters.AddWithValue("@DE_No", recuperer_depot(txtdepot1.Text));
                            cmd.Parameters.AddWithValue("@lsfabrication", Convert.ToDateTime(DateTime.Now));
                            cmd.Parameters.AddWithValue("@user", cbUserCreation.Text);
                            cmd.Parameters.AddWithValue("@lscomplement", txtRef.Text);
                            cmd.Parameters.AddWithValue("@dlnoin", DL_NoIn);

                            cmd.ExecuteNonQuery();


                            cmd.CommandText = @"
                                    ALTER TABLE F_LotSerie CHECK CONSTRAINT ALL; ";
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();

                            //ARTSTOCK
                            cmd.CommandText = @"
                                DISABLE TRIGGER ALL ON F_ARTSTOCK;
                                ALTER TABLE F_ARTSTOCK NOCHECK CONSTRAINT ALL; ";
                            cmd.ExecuteNonQuery();

                            string reference = txtreference.Text;
                            decimal qtes = decimal.Parse(txtqte1.Text.ToString().Replace(",", "."), CultureInfo.InvariantCulture);

                            try
                            {

                                cmd.CommandText = @"
                                         SELECT TOP 1 AS_QteSto
                                         FROM F_ARTSTOCK
                                         WHERE AR_Ref = @AR_Ref AND DE_No = @DE_No";

                                decimal? existingQte = null;

                                cmd.Parameters.AddWithValue("@AR_Ref", reference);
                                cmd.Parameters.AddWithValue("@DE_No", recuperer_depot(txtdepot1.Text));

                                var result = cmd.ExecuteScalar();

                                if (result != null && result != DBNull.Value)
                                {
                                    existingQte = Convert.ToDecimal(result);
                                }

                                if (existingQte == null)
                                {
                                    cmd.CommandText = @"
                                                 INSERT INTO F_ARTSTOCK (
                                                     AR_Ref,
                                                     DE_No,
                                                     DP_NoPrincipal,
                                                     AS_QteSto,
                                                     AS_QteRes,
                                                     AS_QteCom,
                                                     AS_QtePrepa,
                                                     AS_MontSto,
                                                     AS_QteMini,
                                                     AS_QteMaxi
                                                 )
                                                 VALUES (
                                                     @AR_Ref,
                                                     @DE_No,
                                                     @DP_NoPrincipal,
                                                     @Qte,
                                                     0, 0, 0, 0, 0, 0
                                                 )";

                                    cmd.Parameters.Clear();
                                    cmd.Parameters.AddWithValue("@AR_Ref", reference);
                                    cmd.Parameters.AddWithValue("@DE_No", recuperer_depot(txtdepot1.Text));
                                    cmd.Parameters.AddWithValue("@DP_NoPrincipal", 1);
                                    cmd.Parameters.Add("@Qte", SqlDbType.Decimal).Value = qtes;
                                    cmd.Parameters["@Qte"].Precision = 24;
                                    cmd.Parameters["@Qte"].Scale = 6;

                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    cmd.CommandText = @"
                                                 UPDATE F_ARTSTOCK
                                                 SET AS_QteSto = AS_QteSto + @Qte
                                                 WHERE AR_Ref = @AR_Ref AND DE_No = @DE_No";

                                    cmd.Parameters.Clear();
                                    cmd.Parameters.AddWithValue("@AR_Ref", reference);
                                    cmd.Parameters.AddWithValue("@DE_No", recuperer_depot(txtdepot1.Text));
                                    cmd.Parameters.Add("@Qte", SqlDbType.Decimal).Value = qtes;
                                    cmd.Parameters["@Qte"].Precision = 24;
                                    cmd.Parameters["@Qte"].Scale = 6;

                                    cmd.ExecuteNonQuery();
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Erreur");
                            }

                            cmd.CommandText = @"
                                ENABLE TRIGGER ALL ON F_ARTSTOCK;
                                ALTER TABLE F_ARTSTOCK CHECK CONSTRAINT ALL; ";
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"Erreur SQL : {sqlEx.Message}", "Erreur SQL",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }*/
        }

        private void txtnumerolot_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
                e.Handled = true;
        }

        private void txtQteLivrer_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
                e.Handled = true;
        }

        private void frm_qte_livre_Load(object sender, EventArgs e)
        {
            txtLot.Text = "";
            txtnumerolot.Text = "";
        }
    }
}
