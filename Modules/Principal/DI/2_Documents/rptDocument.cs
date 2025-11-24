using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraGrid;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrinting.Preview;
using DevExpress.XtraReports.UI;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using static System.Data.Entity.Infrastructure.Design.Executor;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class rptDocument : DevExpress.XtraReports.UI.XtraReport
    {
        public rptDocument()
        {
            InitializeComponent();
           
        }
        public static DateTime doDate;
        public static DateTime doDateLivrPrev;
        public static int doImprim;
        public static string doTiers;
        public static int doStatut;
        public static int doReliquat;
        public static int deno;
        public static string doRef;
        public static int TypeAchat;
        public static int doExpedit;
        public static string doPiece;
        public static int doTaxe1;
        public static string doCodeTaxe1;
        public static int? CoNo;
        private readonly F_COLLABORATEURRepository _collaborateurRepository;
        public static string doEntete;
        public static string a_type;
        public static DataTable dt ;

        private void rptDocument_ParametersRequestSubmit(object sender, DevExpress.XtraReports.Parameters.ParametersRequestEventArgs e)
        {
            string doPiece = this.Parameters["pDocumentNumber"].Value.ToString();

            DataTable enteteTable = GetEnteteData(doPiece);

            if (enteteTable.Rows.Count > 0)
            {
                DataRow row = enteteTable.Rows[0];

                // LES ENTETES
                xrLabel1.Text = row["DO_Piece"].ToString();
                xrLabel4.Text = Convert.ToDateTime(row["DO_Date"]).ToString("dd/MM/yyyy");
                xrLabel6.Text = row["CT_Intitule"].ToString();
                xrLabel13.Text = row["A_TYPE"].ToString();
                xrLabel15.Text = row["Acheteur"].ToString();
                // LES PIEDS 
                decimal DO_TotalHT = Convert.ToDecimal(row["DO_TotalHT"]);
                decimal DO_TotalTTC = Convert.ToDecimal(row["DO_TotalTTC"]);

                //xrLabel7.Text = DO_TotalHT.ToString("N2");
                //xrLabel10.Text = DO_TotalTTC.ToString("N2");

                //double montant = Convert.ToDouble(DO_TotalTTC);
                
                
            }
            DataTable ligneTable = ChargerLignes(doPiece);
            DetailReport.DataSource = ligneTable;
            DataTable InfosTable = GetFraisImport(doPiece);
            // ON AFFICHE LES LIGNES InfosTable ICI
           
            AjouterFraisDansReportFooter(this.GroupFooter1, InfosTable);

            // CONVERTIR EN LETTRE

            string partInt = Converter.ConvertNumberToWords(totalFrais, Language.French);
            int decimales = (int)((totalFrais - Math.Floor(totalFrais)) * 100);
            string partDec = decimales == 0 ? "" : " " + Converter.AjoutDecimal(totalFrais, Language.French);
            xrLabel9.Text = partInt + " " + partDec;
        }
        private decimal totalFrais;
        private void AjouterFraisDansReportFooter(GroupFooterBand footerBand, DataTable table)
        {
            string[] colonnesAAfficher = { "FR_Intitule", "FI_Montant" };
            footerBand.Controls.Clear();

            XRTable xrTable = new XRTable
            {
                Borders = DevExpress.XtraPrinting.BorderSide.All,
                Font = new Font("Segoe UI", 9F),
                WidthF = this.PageWidth - this.Margins.Left - this.Margins.Right
            };

            xrTable.BeginInit();

            // En-tête
            XRTableRow headerRow = new XRTableRow();
            foreach (string colName in colonnesAAfficher)
            {
                XRTableCell headerCell = new XRTableCell
                {
                    Text = colName,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                    BackColor = Color.LightGray,
                    Padding = new DevExpress.XtraPrinting.PaddingInfo(5, 5, 5, 5)
                };
                headerRow.Cells.Add(headerCell);
            }
            xrTable.Rows.Add(headerRow);

            // Corps du tableau
            totalFrais = 0;

            foreach (DataRow dataRow in table.Rows)
            {
                XRTableRow row = new XRTableRow();
                foreach (string colName in colonnesAAfficher)
                {
                    object val = dataRow[colName];
                    string text = val != DBNull.Value ? val.ToString() : "";

                    // Calcul total
                    if (colName == "FI_Montant" && val != DBNull.Value)
                    {
                        totalFrais += Convert.ToDecimal(val);
                    }

                    XRTableCell cell = new XRTableCell
                    {
                        TextAlignment = (colName == "FI_Montant") ? TextAlignment.MiddleRight : TextAlignment.MiddleLeft,
                        Padding = new PaddingInfo(5, 5, 5, 5)
                    };

                    if (colName == "FI_Montant" && val != DBNull.Value)
                    {
                        decimal montant = Convert.ToDecimal(val);
                        cell.Text = montant.ToString("N2"); // Format avec 2 décimales
                    }
                    else
                    {
                        cell.Text = text;
                    }


                    row.Cells.Add(cell);
                }
                xrTable.Rows.Add(row);
                xrTable.HeightF= 25f;
            }

            // Ligne du total
            XRTableRow totalRow = new XRTableRow();
            XRTableCell libelleCell = new XRTableCell
            {
                Text = "Total",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight,
                Padding = new DevExpress.XtraPrinting.PaddingInfo(5, 5, 5, 5)
            };
            totalRow.Cells.Add(libelleCell);

            XRTableCell montantCell = new XRTableCell
            {
                Text = totalFrais.ToString("N2"),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight,
                Padding = new DevExpress.XtraPrinting.PaddingInfo(5, 5, 5, 5)
            };
            totalRow.Cells.Add(montantCell);

            xrTable.Rows.Add(totalRow);

            xrTable.EndInit();
            xrTable.LocationF = new PointF(0, 0);
            footerBand.Controls.Add(xrTable);
        }


        private static string dbPrincipale = ucDocuments.dbNamePrincipale;
        private static string serveripPrincipale = ucDocuments.serverIpPrincipale;
        public static int rownum = 0;
        private static DataTable dataTable;
        private static string connectionString = $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";
        private static SqlDataAdapter dataAdapter;
        private DataTable GetEnteteData(string dopiece)
        {
            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = $"SELECT * FROM dbo.ACHAT_ENTETE WHERE DO_Piece = @DOPiece";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@DOPiece", dopiece);
                    dataAdapter = new SqlDataAdapter(cmd); 
                    SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                    dataTable = new DataTable();
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                    
                    conn.Open();
                    dataAdapter.Fill(dt);
                }
            }
            return dt;
        }
        private DataTable ChargerLignes(string dopiecetxt)
        {
            
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                string query = $@"SELECT dbo.F_COMPTET.CT_Intitule AS CT_Intitule, *, 'Remove' AS Action, 'Update' AS Validation, 'Add' AS Insertion 
                                    FROM dbo.F_DOCLIGNE 
                                    INNER JOIN dbo.F_COMPTET ON (dbo.F_DOCLIGNE.CT_Num = dbo.F_COMPTET.CT_Num)
                                    WHERE DO_Piece = @dopiecetxt AND Retenu = 1";

                using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@dopiecetxt", dopiecetxt);
                        dataAdapter = new SqlDataAdapter(cmd); // On passe `cmd`, pas `query`
                        SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                        dataTable = new DataTable();

                        connection.Open();
                        dataAdapter.Fill(dataTable);
                        return dataTable;
                    }
                }

            
        }

        private void ShowPreviewWithTracking(XtraReport report, string docPiece)
        {
            ReportPrintTool printTool = new ReportPrintTool(report);
            PrintPreviewFormEx previewForm = printTool.PreviewForm as PrintPreviewFormEx;

            // Handle the "Print" click
            report.PrintingSystem.StartPrint += (s, e) =>
            {
                try
                {
                    UpdateImprimFlag(docPiece);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur lors de la mise à jour de l'impression : " + ex.Message);
                }
            };
                        
            previewForm.ShowDialog();
        }
        private void UpdateImprimFlag(string docPiece)
        {
            using (var context = new AppDbContext())
            {
                var doc = context.F_DOCENTETE.FirstOrDefault(d => d.DO_Piece == docPiece);
                if (doc != null)
                {
                    doc.DO_Imprim = 1;
                    context.SaveChanges();
                }
            }
        }
        private static DataTable dataTableFI;
        private DataTable GetFraisImport(string dopiece)
        {
            DataTable dtFrais = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = $@"
                                SELECT 
                                  dbo.F_DOCENTETE.DO_Piece,
                                  'DO_TotalHT' AS FR_Intitule,
                                  dbo.F_DOCENTETE.DO_TotalHT AS FI_Montant,
                                  '' AS FI_Devise      
                                FROM
                                  dbo.F_DOCENTETE
                                    WHERE DO_Piece = @DOPiece
                                    /*UNION ALL
                                SELECT 
                                  dbo.F_DOCENTETE.DO_Piece,
                                  'DO_TotalTTC' AS FR_Intitule,
                                  dbo.F_DOCENTETE.DO_TotalTTC FI_Montant,
                                  '' AS FI_Devise      
                                FROM
                                  dbo.F_DOCENTETE
                                    WHERE DO_Piece = @DOPiece*/
                                    UNION ALL   
                                    SELECT 
                                  dbo.F_DOCFRAISIMPORT.DO_Piece,   
                                  dbo.P_TYPEFRAIS.FR_Intitule,
                                  dbo.F_DOCFRAISIMPORT.FI_Montant,
                                  dbo.F_DOCFRAISIMPORT.FI_Devise
                                FROM
                                  dbo.F_DOCFRAISIMPORT
                                  INNER JOIN dbo.P_TYPEFRAIS ON (dbo.F_DOCFRAISIMPORT.FI_TypeFraisId = dbo.P_TYPEFRAIS.FR_Num) 
                                WHERE DO_Piece = @DOPiece";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@DOPiece", dopiece);
                    dataAdapter = new SqlDataAdapter(cmd);
                    SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                    dataTableFI = new DataTable();
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }

                    conn.Open();
                    dataAdapter.Fill(dataTableFI);
                }
            }
            return dataTableFI;
        }

    }
}
