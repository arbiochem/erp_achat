using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI
{
    public class Lignes
    {
        public static int rownum = 0;
        private static SqlConnection connection;
        private static DataTable dataTable;
        private static string dbPrincipale = ucDocuments.dbNamePrincipale;
        private static string serveripPrincipale = ucDocuments.serverIpPrincipale;
        private static string connectionString = $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";
        private static SqlDataAdapter dataAdapter;
        private static string DOPiece;
        public static void AfficherLignes(GridControl gc, string dopiece)
        {
            try
            {
                DOPiece = dopiece;
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    string query = $@"SELECT 
                                              dbo.F_COMPTET.CT_Intitule,
                                              dbo.F_DOCLIGNE.DO_Domaine,
                                              dbo.F_DOCLIGNE.DO_Type,
                                              dbo.F_DOCLIGNE.CT_Num,
                                              dbo.F_DOCLIGNE.cbCT_Num,
                                              dbo.F_DOCLIGNE.DO_Piece,
                                              dbo.F_DOCLIGNE.cbDO_Piece,
                                              dbo.F_DOCLIGNE.DL_PieceBC,
                                              dbo.F_DOCLIGNE.cbDL_PieceBC,
                                              dbo.F_DOCLIGNE.DL_PieceBL,
                                              dbo.F_DOCLIGNE.cbDL_PieceBL,
                                              dbo.F_DOCLIGNE.DO_Date,
                                              dbo.F_DOCLIGNE.DL_DateBC,
                                              dbo.F_DOCLIGNE.DL_DateBL,
                                              dbo.F_DOCLIGNE.DL_Ligne,
                                              dbo.F_DOCLIGNE.DO_Ref,
                                              dbo.F_DOCLIGNE.cbDO_Ref,
                                              dbo.F_DOCLIGNE.DL_TNomencl,
                                              dbo.F_DOCLIGNE.DL_TRemPied,
                                              dbo.F_DOCLIGNE.DL_TRemExep,
                                              dbo.F_DOCLIGNE.AR_Ref,
                                              dbo.F_DOCLIGNE.cbAR_Ref,
                                              dbo.F_DOCLIGNE.DL_Design,
                                              dbo.F_DOCLIGNE.DL_PrixUnitaire,
                                              dbo.F_DOCLIGNE.DL_Remise01REM_Valeur,
                                              dbo.F_DOCLIGNE.DL_Qte,
                                              dbo.F_DOCLIGNE.DL_QteBC,
                                              dbo.F_DOCLIGNE.DL_QteBL,
                                              dbo.F_DOCLIGNE.DL_PoidsNet,
                                              dbo.F_DOCLIGNE.DL_PoidsBrut,
                                              dbo.F_DOCLIGNE.DL_Remise01REM_Type,
                                              dbo.F_DOCLIGNE.DL_Remise02REM_Valeur,
                                              dbo.F_DOCLIGNE.DL_Remise02REM_Type,
                                              dbo.F_DOCLIGNE.DL_Remise03REM_Valeur,
                                              dbo.F_DOCLIGNE.DL_Remise03REM_Type,                                             
                                              dbo.F_DOCLIGNE.DL_PUBC,
                                              dbo.F_DOCLIGNE.DL_Taxe1,
                                              dbo.F_DOCLIGNE.DL_TypeTaux1,
                                              dbo.F_DOCLIGNE.DL_TypeTaxe1,
                                              dbo.F_DOCLIGNE.DL_Taxe2,
                                              dbo.F_DOCLIGNE.DL_TypeTaux2,
                                              dbo.F_DOCLIGNE.DL_TypeTaxe2,
                                              dbo.F_DOCLIGNE.CO_No,
                                              dbo.F_DOCLIGNE.cbCO_No,
                                              dbo.F_DOCLIGNE.AG_No1,
                                              dbo.F_DOCLIGNE.AG_No2,
                                              dbo.F_DOCLIGNE.DL_PrixRU,
                                              dbo.F_DOCLIGNE.DL_CMUP,
                                              dbo.F_DOCLIGNE.DL_MvtStock,
                                              dbo.F_DOCLIGNE.DT_No,
                                              dbo.F_DOCLIGNE.cbDT_No,
                                              dbo.F_DOCLIGNE.AF_RefFourniss,
                                              dbo.F_DOCLIGNE.cbAF_RefFourniss,
                                              dbo.F_DOCLIGNE.EU_Enumere,
                                              dbo.F_DOCLIGNE.EU_Qte,
                                              dbo.F_DOCLIGNE.DL_TTC,
                                              dbo.F_DOCLIGNE.DE_No,
                                              dbo.F_DOCLIGNE.cbDE_No,
                                              dbo.F_DOCLIGNE.DL_NoRef,
                                              dbo.F_DOCLIGNE.DL_TypePL,
                                              dbo.F_DOCLIGNE.DL_PUDevise,
                                              dbo.F_DOCLIGNE.DL_PUTTC,
                                              dbo.F_DOCLIGNE.DL_No,
                                              dbo.F_DOCLIGNE.DO_DateLivr,
                                              dbo.F_DOCLIGNE.CA_Num,
                                              dbo.F_DOCLIGNE.cbCA_Num,
                                              dbo.F_DOCLIGNE.DL_Taxe3,
                                              dbo.F_DOCLIGNE.DL_TypeTaux3,
                                              dbo.F_DOCLIGNE.DL_TypeTaxe3,
                                              dbo.F_DOCLIGNE.DL_Frais,
                                              dbo.F_DOCLIGNE.DL_Valorise,
                                              dbo.F_DOCLIGNE.AR_RefCompose,
                                              dbo.F_DOCLIGNE.cbAR_RefCompose,
                                              dbo.F_DOCLIGNE.DL_NonLivre,
                                              dbo.F_DOCLIGNE.AC_RefClient,
                                              dbo.F_DOCLIGNE.DL_MontantHT,
                                              dbo.F_DOCLIGNE.DL_MontantTTC,
                                              dbo.F_DOCLIGNE.DL_FactPoids,
                                              dbo.F_DOCLIGNE.DL_Escompte,
                                              dbo.F_DOCLIGNE.DL_PiecePL,
                                              dbo.F_DOCLIGNE.cbDL_PiecePL,
                                              dbo.F_DOCLIGNE.DL_DatePL,
                                              dbo.F_DOCLIGNE.DL_QtePL,
                                              dbo.F_DOCLIGNE.DL_NoColis,
                                              dbo.F_DOCLIGNE.DL_NoLink,
                                              dbo.F_DOCLIGNE.cbDL_NoLink,
                                              dbo.F_DOCLIGNE.RP_Code,
                                              dbo.F_DOCLIGNE.cbRP_Code,
                                              dbo.F_DOCLIGNE.DL_QteRessource,
                                              dbo.F_DOCLIGNE.DL_DateAvancement,
                                              dbo.F_DOCLIGNE.PF_Num,
                                              dbo.F_DOCLIGNE.cbPF_Num,
                                              dbo.F_DOCLIGNE.DL_CodeTaxe1,
                                              dbo.F_DOCLIGNE.DL_CodeTaxe2,
                                              dbo.F_DOCLIGNE.DL_CodeTaxe3,
                                              dbo.F_DOCLIGNE.DL_PieceOFProd,
                                              dbo.F_DOCLIGNE.DL_PieceDE,
                                              dbo.F_DOCLIGNE.cbDL_PieceDE,
                                              dbo.F_DOCLIGNE.DL_DateDE,
                                              dbo.F_DOCLIGNE.DL_QteDE,
                                              dbo.F_DOCLIGNE.DL_Operation,
                                              dbo.F_DOCLIGNE.DL_NoSousTotal,
                                              dbo.F_DOCLIGNE.CA_No,
                                              dbo.F_DOCLIGNE.cbCA_No,
                                              dbo.F_DOCLIGNE.DO_DocType,
                                              dbo.F_DOCLIGNE.cbProt,
                                              dbo.F_DOCLIGNE.cbMarq,
                                              dbo.F_DOCLIGNE.cbCreateur,
                                              dbo.F_DOCLIGNE.cbModification,
                                              dbo.F_DOCLIGNE.cbReplication,
                                              dbo.F_DOCLIGNE.cbFlag,
                                              dbo.F_DOCLIGNE.cbCreation,
                                              dbo.F_DOCLIGNE.cbCreationUser,
                                              dbo.F_DOCLIGNE.cbHash,
                                              dbo.F_DOCLIGNE.cbHashVersion,
                                              dbo.F_DOCLIGNE.cbHashDate,
                                              dbo.F_DOCLIGNE.cbHashOrder,
                                              dbo.F_DOCLIGNE.DL_MontantRegle,
                                              dbo.F_DOCLIGNE.DL_PieceFourniss,
                                              dbo.F_DOCLIGNE.DL_DatePieceFourniss,

                                              dbo.F_DOCLIGNE.Retenu,
                                              'Remove' AS Action,
                                              'Update' AS Validation
                                                
                                            FROM
                                              dbo.F_DOCLIGNE
                                              INNER JOIN dbo.F_COMPTET ON (dbo.F_DOCLIGNE.CT_Num = dbo.F_COMPTET.CT_Num)
                                            WHERE DO_Domaine = 1 AND DO_Piece = @DO_Piece
                                            ORDER BY DL_Ligne ASC";
                    
                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@DO_Piece", dopiece);
                        dataAdapter = new SqlDataAdapter(cmd); // On passe `cmd`, pas `query`
                        SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                        dataTable = new DataTable();

                        connection.Open();
                        dataAdapter.Fill(dataTable);
                        rownum = dataTable.Rows.Count;
                        gc.DataSource = dataTable;

                    }
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static void cacherColonnes(GridView view)
        {       
            view.Columns["CT_Intitule"].VisibleIndex = -1;
            view.Columns["DO_Domaine"].VisibleIndex = -1;
            view.Columns["DO_Type"].VisibleIndex = -1;
            view.Columns["cbCT_Num"].VisibleIndex = -1;
            view.Columns["DO_Piece"].VisibleIndex = -1;
            view.Columns["cbDO_Piece"].VisibleIndex = -1;
            view.Columns["DL_PieceBC"].VisibleIndex = -1;
            view.Columns["cbDL_PieceBC"].VisibleIndex = -1;
            view.Columns["DL_PieceBL"].VisibleIndex = -1;
            view.Columns["cbDL_PieceBL"].VisibleIndex = -1;
            view.Columns["DO_Date"].VisibleIndex = -1;
            view.Columns["DL_DateBC"].VisibleIndex = -1;
            view.Columns["DL_DateBL"].VisibleIndex = -1;
            view.Columns["DL_Ligne"].VisibleIndex = -1;
            view.Columns["DO_Ref"].VisibleIndex = -1;
            view.Columns["cbDO_Ref"].VisibleIndex = -1;
            view.Columns["DL_TNomencl"].VisibleIndex = -1;
            view.Columns["DL_TRemPied"].VisibleIndex = -1;
            view.Columns["DL_TRemExep"].VisibleIndex = -1;
            //view.Columns["AR_Ref"].VisibleIndex = -1;
            //view.Columns["DL_Design"].VisibleIndex = -1;
            view.Columns["cbAR_Ref"].VisibleIndex = -1;
            //view.Columns["DL_Qte"].VisibleIndex = -1;
            view.Columns["DL_QteBC"].VisibleIndex = -1;
            view.Columns["DL_QteBL"].VisibleIndex = -1;
            //view.Columns["DL_PoidsNet"].VisibleIndex = -1;
            view.Columns["DL_Valorise"].VisibleIndex = -1;
            view.Columns["DL_PoidsBrut"].VisibleIndex = -1;
            view.Columns["DL_Remise01REM_Valeur"].VisibleIndex = -1;
            view.Columns["DL_Remise01REM_Type"].VisibleIndex = -1;
            view.Columns["DL_Remise02REM_Valeur"].VisibleIndex = -1;
            view.Columns["DL_Remise02REM_Type"].VisibleIndex = -1;
            view.Columns["DL_Remise03REM_Valeur"].VisibleIndex = -1;
            view.Columns["DL_Remise03REM_Type"].VisibleIndex = -1;
            view.Columns["DL_PUBC"].VisibleIndex = -1;
            //view.Columns["DL_PrixUnitaire"].VisibleIndex = -1;
            view.Columns["DL_PUBC"].VisibleIndex = -1;
            view.Columns["DL_Taxe1"].VisibleIndex = -1;
            view.Columns["DL_TypeTaux1"].VisibleIndex = -1;
            view.Columns["DL_TypeTaxe1"].VisibleIndex = -1;
            view.Columns["DL_Taxe2"].VisibleIndex = -1;
            view.Columns["DL_TypeTaux2"].VisibleIndex = -1;
            view.Columns["DL_TypeTaxe2"].VisibleIndex = -1;
            view.Columns["CO_No"].VisibleIndex = -1;
            view.Columns["cbCO_No"].VisibleIndex = -1;
            view.Columns["AG_No1"].VisibleIndex = -1;
            view.Columns["AG_No2"].VisibleIndex = -1;
            //view.Columns["DL_PrixRU"].VisibleIndex = -1;
            view.Columns["DL_CMUP"].VisibleIndex = -1;
            view.Columns["DL_MvtStock"].VisibleIndex = -1;
            view.Columns["DT_No"].VisibleIndex = -1;
            view.Columns["cbDT_No"].VisibleIndex = -1;
            view.Columns["cbAF_RefFourniss"].VisibleIndex = -1;
            view.Columns["EU_Enumere"].VisibleIndex = -1;
            view.Columns["EU_Qte"].VisibleIndex = -1;
            view.Columns["DL_TTC"].VisibleIndex = -1;
            view.Columns["DE_No"].VisibleIndex = -1;
            view.Columns["cbDE_No"].VisibleIndex = -1;
            view.Columns["DL_NoRef"].VisibleIndex = -1;
            view.Columns["DL_TypePL"].VisibleIndex = -1;
            view.Columns["DL_PUDevise"].VisibleIndex = -1;
            view.Columns["DL_PUTTC"].VisibleIndex = -1;
            view.Columns["DL_No"].VisibleIndex = -1;
            view.Columns["DO_DateLivr"].VisibleIndex = -1;
            view.Columns["CA_Num"].VisibleIndex = -1;
            view.Columns["cbCA_Num"].VisibleIndex = -1;
            view.Columns["DL_Taxe3"].VisibleIndex = -1;
            view.Columns["DL_TypeTaux3"].VisibleIndex = -1;
            view.Columns["DL_TypeTaxe3"].VisibleIndex = -1;
            view.Columns["AR_RefCompose"].VisibleIndex = -1;
            view.Columns["cbAR_RefCompose"].VisibleIndex = -1;
            view.Columns["AC_RefClient"].VisibleIndex = -1;
            //view.Columns["DL_MontantHT"].VisibleIndex = -1;
            view.Columns["DL_MontantTTC"].VisibleIndex = -1;
            view.Columns["DL_FactPoids"].VisibleIndex = -1;
            view.Columns["DL_Escompte"].VisibleIndex = -1;
            view.Columns["DL_PiecePL"].VisibleIndex = -1;
            view.Columns["cbDL_PiecePL"].VisibleIndex = -1;
            view.Columns["DL_DatePL"].VisibleIndex = -1;
            view.Columns["DL_QtePL"].VisibleIndex = -1;
            view.Columns["DL_NoColis"].VisibleIndex = -1;
            view.Columns["DL_NoLink"].VisibleIndex = -1;
            view.Columns["cbDL_NoLink"].VisibleIndex = -1;
            view.Columns["RP_Code"].VisibleIndex = -1;
            view.Columns["cbRP_Code"].VisibleIndex = -1;
            view.Columns["DL_QteRessource"].VisibleIndex = -1;
            view.Columns["DL_DateAvancement"].VisibleIndex = -1;
            view.Columns["PF_Num"].VisibleIndex = -1;
            view.Columns["cbPF_Num"].VisibleIndex = -1;
            view.Columns["DL_CodeTaxe1"].VisibleIndex = -1;
            view.Columns["DL_CodeTaxe2"].VisibleIndex = -1;
            view.Columns["DL_CodeTaxe3"].VisibleIndex = -1;
            view.Columns["DL_PieceOFProd"].VisibleIndex = -1;
            view.Columns["DL_PieceDE"].VisibleIndex = -1;
            view.Columns["cbDL_PieceDE"].VisibleIndex = -1;
            view.Columns["DL_DateDE"].VisibleIndex = -1;
            view.Columns["DL_QteDE"].VisibleIndex = -1;
            view.Columns["DL_Operation"].VisibleIndex = -1;
            view.Columns["DL_NoSousTotal"].VisibleIndex = -1;
            view.Columns["CA_No"].VisibleIndex = -1;
            view.Columns["cbCA_No"].VisibleIndex = -1;
            view.Columns["DO_DocType"].VisibleIndex = -1;
            view.Columns["cbProt"].VisibleIndex = -1;
            view.Columns["cbMarq"].VisibleIndex = -1;
            view.Columns["cbCreateur"].VisibleIndex = -1;
            view.Columns["cbModification"].VisibleIndex = -1;
            view.Columns["cbReplication"].VisibleIndex = -1;
            view.Columns["cbFlag"].VisibleIndex = -1;
            view.Columns["cbCreation"].VisibleIndex = -1;
            view.Columns["cbCreationUser"].VisibleIndex = -1;
            view.Columns["cbHash"].VisibleIndex = -1;
            view.Columns["cbHashVersion"].VisibleIndex = -1;
            view.Columns["cbHashDate"].VisibleIndex = -1;
            view.Columns["cbHashOrder"].VisibleIndex = -1;
            if (DOPiece!= "" && DOPiece.Substring(0, 3) == "APA")
            {
                view.Columns["DL_MontantRegle"].VisibleIndex = -1;
            }
            else
            {
                view.Columns["DL_MontantRegle"].VisibleIndex = 10;
            }

        }
        public static DataTable CreateTableLigne()
        {
            DataTable tbl = new DataTable();

            tbl.Columns.Add("CT_Intitule", typeof(string));
            tbl.Columns.Add("DO_Domaine", typeof(int));
            tbl.Columns.Add("DO_Type", typeof(int));
            tbl.Columns.Add("CT_Num", typeof(string));
            tbl.Columns.Add("cbCT_Num", typeof(string));
            tbl.Columns.Add("DO_Piece", typeof(string));
            tbl.Columns.Add("cbDO_Piece", typeof(string));
            tbl.Columns.Add("DL_PieceBC", typeof(string));
            tbl.Columns.Add("cbDL_PieceBC", typeof(string));
            tbl.Columns.Add("DL_PieceBL", typeof(string));
            tbl.Columns.Add("cbDL_PieceBL", typeof(string));
            tbl.Columns.Add("DO_Date", typeof(DateTime));
            tbl.Columns.Add("DL_DateBC", typeof(DateTime));
            tbl.Columns.Add("DL_DateBL", typeof(DateTime));
            tbl.Columns.Add("DL_Ligne", typeof(int));
            tbl.Columns.Add("DO_Ref", typeof(string));
            tbl.Columns.Add("cbDO_Ref", typeof(string));
            tbl.Columns.Add("DL_TNomencl", typeof(int));
            tbl.Columns.Add("DL_TRemPied", typeof(int));
            tbl.Columns.Add("DL_TRemExep", typeof(int));
            tbl.Columns.Add("AR_Ref", typeof(string));
            tbl.Columns.Add("cbAR_Ref", typeof(string));
            tbl.Columns.Add("DL_Design", typeof(string));
            tbl.Columns.Add("DL_Qte", typeof(decimal));
            tbl.Columns.Add("DL_QteBC", typeof(decimal));
            tbl.Columns.Add("DL_QteBL", typeof(decimal));
            tbl.Columns.Add("DL_PoidsNet", typeof(decimal));
            tbl.Columns.Add("DL_Valorise", typeof(decimal));
            tbl.Columns.Add("DL_PoidsBrut", typeof(double));
            tbl.Columns.Add("DL_Remise01REM_Valeur", typeof(int));
            tbl.Columns.Add("DL_Remise01REM_Type", typeof(int));
            tbl.Columns.Add("DL_Remise02REM_Valeur", typeof(int));
            tbl.Columns.Add("DL_Remise02REM_Type", typeof(int));
            tbl.Columns.Add("DL_Remise03REM_Valeur", typeof(int));
            tbl.Columns.Add("DL_Remise03REM_Type", typeof(int));
            tbl.Columns.Add("DL_PUBC", typeof(decimal));
            tbl.Columns.Add("DL_PrixUnitaire", typeof(decimal));
            tbl.Columns.Add("DL_Taxe1", typeof(decimal));
            tbl.Columns.Add("DL_TypeTaux1", typeof(int));
            tbl.Columns.Add("DL_TypeTaxe1", typeof(int));
            tbl.Columns.Add("DL_Taxe2", typeof(int));
            tbl.Columns.Add("DL_TypeTaux2", typeof(int));
            tbl.Columns.Add("DL_TypeTaxe2", typeof(int));
            tbl.Columns.Add("CO_No", typeof(int));
            tbl.Columns.Add("cbCO_No", typeof(int));
            tbl.Columns.Add("AG_No1", typeof(int));
            tbl.Columns.Add("AG_No2", typeof(int));
            tbl.Columns.Add("DL_PrixRU", typeof(decimal));
            tbl.Columns.Add("DL_CMUP", typeof(decimal));
            tbl.Columns.Add("DL_MvtStock", typeof(int));
            tbl.Columns.Add("DT_No", typeof(int));
            tbl.Columns.Add("cbDT_No", typeof(int));
            tbl.Columns.Add("cbAF_RefFourniss", typeof(string));
            tbl.Columns.Add("EU_Enumere", typeof(string));
            tbl.Columns.Add("EU_Qte", typeof(decimal));
            tbl.Columns.Add("DL_TTC", typeof(int));
            tbl.Columns.Add("DE_No", typeof(int));
            tbl.Columns.Add("cbDE_No", typeof(int));
            tbl.Columns.Add("DL_NoRef", typeof(int));
            tbl.Columns.Add("DL_TypePL", typeof(int));
            tbl.Columns.Add("DL_PUDevise", typeof(decimal));
            tbl.Columns.Add("DL_PUTTC", typeof(decimal));
            tbl.Columns.Add("DL_No", typeof(int));
            tbl.Columns.Add("DO_DateLivr", typeof(DateTime));
            tbl.Columns.Add("CA_Num", typeof(string));
            tbl.Columns.Add("cbCA_Num", typeof(string));
            tbl.Columns.Add("DL_Taxe3", typeof(decimal));
            tbl.Columns.Add("DL_TypeTaux3", typeof(int));
            tbl.Columns.Add("DL_TypeTaxe3", typeof(int));
            tbl.Columns.Add("cbAR_RefCompose", typeof(string));
            tbl.Columns.Add("AC_RefClient", typeof(string));
            tbl.Columns.Add("DL_MontantHT", typeof(decimal));
            tbl.Columns.Add("DL_MontantTTC", typeof(decimal));
            tbl.Columns.Add("DL_FactPoids", typeof(int));
            tbl.Columns.Add("DL_Escompte", typeof(int));
            tbl.Columns.Add("DL_PiecePL", typeof(string));
            tbl.Columns.Add("cbDL_PiecePL", typeof(string));
            tbl.Columns.Add("DL_DatePL", typeof(DateTime));
            tbl.Columns.Add("DL_QtePL", typeof(decimal));
            tbl.Columns.Add("DL_NoColis", typeof(string));
            tbl.Columns.Add("DL_NoLink", typeof(int));
            tbl.Columns.Add("cbDL_NoLink", typeof(int));
            tbl.Columns.Add("RP_Code", typeof(string));
            tbl.Columns.Add("cbRP_Code", typeof(string));
            tbl.Columns.Add("DL_QteRessource", typeof(int));
            tbl.Columns.Add("DL_DateAvancement", typeof(DateTime));
            tbl.Columns.Add("PF_Num", typeof(string));
            tbl.Columns.Add("cbPF_Num", typeof(string));
            tbl.Columns.Add("DL_CodeTaxe1", typeof(string));
            tbl.Columns.Add("DL_CodeTaxe2", typeof(string));
            tbl.Columns.Add("DL_CodeTaxe3", typeof(string));
            tbl.Columns.Add("DL_PieceOFProd", typeof(int));
            tbl.Columns.Add("DL_PieceDE", typeof(string));
            tbl.Columns.Add("cbDL_PieceDE", typeof(string));
            tbl.Columns.Add("DL_DateDE", typeof(DateTime));
            tbl.Columns.Add("DL_QteDE", typeof(decimal));
            tbl.Columns.Add("DL_MontantRegle", typeof(decimal));
            tbl.Columns.Add("DL_Operation", typeof(string));
            tbl.Columns.Add("DL_NoSousTotal", typeof(int));
            tbl.Columns.Add("CA_No", typeof(int));
            tbl.Columns.Add("cbCA_No", typeof(int));
            tbl.Columns.Add("DO_DocType", typeof(int));
            tbl.Columns.Add("cbProt", typeof(int));
            tbl.Columns.Add("cbMarq", typeof(int));
            tbl.Columns.Add("cbCreateur", typeof(string));
            tbl.Columns.Add("cbModification", typeof(DateTime));
            tbl.Columns.Add("cbReplication", typeof(int));
            tbl.Columns.Add("cbFlag", typeof(int));
            tbl.Columns.Add("cbCreation", typeof(DateTime));
            tbl.Columns.Add("cbCreationUser", typeof(string));
            tbl.Columns.Add("cbHash", typeof(string));
            tbl.Columns.Add("cbHashVersion", typeof(int));
            tbl.Columns.Add("cbHashDate", typeof(DateTime));
            tbl.Columns.Add("cbHashOrder", typeof(int));

            tbl.Columns.Add("DL_PieceFourniss", typeof(string));
            tbl.Columns.Add("DL_DatePieceFourniss", typeof(string));

            tbl.Columns.Add("Action", typeof(string));
            tbl.Columns.Add("Retenu", typeof(byte));
            tbl.Columns.Add("Validation", typeof(string));

            

            return tbl;
        }
    }
}
