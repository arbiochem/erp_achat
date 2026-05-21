using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Models;
using DevExpress.CodeParser;
using DevExpress.DataAccess.Native.Excel;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
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
    public class Entetes
    {
        public static int rownum = 0;
        private static SqlConnection connection;
        private static DataTable dataTable;
        private static DataTable dataTableFrns;
        //private static string dbname = ucDocuments.dbNamePrincipale;

        private static string DbPrincipale => ucDocuments.dbNamePrincipale;
        private static string ServerIpPrincipale => ucDocuments.serverIpPrincipale;


        private static string connectionString = $"Server={ServerIpPrincipale};Database={DbPrincipale};" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";

        private static string connectionStrings = $"Server={ServerIpPrincipale};" +
                                $"Database=TRANSIT;User ID=Dev;Password=1234;" +
                                $"TrustServerCertificate=True;Connection Timeout=120;";
        private static SqlDataAdapter dataAdapter;

        public static void AfficherEntetes(GridControl gc, int achattype, BindingSource bs)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    string query = "";
                    /*if(achattype != 200)
                    {
                        query =  $"SELECT * FROM dbo.ACHAT_ENTETE WHERE DO_Type = @dotype ORDER BY DO_Date DESC";
                    }
                    else
                    {*/
                        query = $"SELECT * FROM dbo.ACHAT_ENTETE ORDER BY DO_Date DESC";
                    //}
                        try
                        {
                            using (SqlCommand cmd = new SqlCommand(query, connection))
                            {
                                // cmd.Parameters.AddWithValue("@dotype", achattype);

                                if (connection.State != ConnectionState.Open)
                                    connection.Open();

                                using (SqlDataAdapter localAdapter = new SqlDataAdapter(cmd))
                                {
                                    dataTable = new DataTable();
                                    localAdapter.Fill(dataTable);

                                    rownum = dataTable.Rows.Count;

                                    if (rownum == 0)
                                    {
                                        bs.DataSource = null;
                                        gc.DataSource = null;
                                    }
                                    else
                                    {
                                        bs.DataSource = dataTable;
                                        gc.DataSource = bs;
                                    }
                                }
                            }
                        }
                        catch (SqlException ex)
                        {
                            MessageBox.Show("Erreur SQL : " + ex.Message);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Erreur : " + ex.Message);
                        }
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void AfficherEntetes_achat(GridControl gc, GridControl gc1, GridControl gc2, GridControl gc3, int achattype, BindingSource bs)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    // Vérifier si DO_Cloture existe dans la table
                    string query = @"SELECT DO_Piece,CT_Intitule, DO_TotalTTC,
                            CASE WHEN COL_LENGTH('dbo.ACHAT_ENTETE', 'DO_Cloture') IS NOT NULL 
                                 THEN CAST(DO_Cloture AS INT) 
                                 ELSE 0 
                            END AS DO_Cloture,CO_No,DE_No,DO_TIERS,DO_Statut,DO_Expedit,DO_Coord01,DO_Ref,DO_DateLivr,DO_Date,DO_DateExpedition,A_TYPE,DO_CodeTaxe1,DO_Taxe1,DO_Type,DO_Imprim,DO_Reliquat
                            FROM dbo.ACHAT_ENTETE ORDER BY DO_Piece ASC";
                    try
                    {
                        using (SqlCommand cmd = new SqlCommand(query, connection))
                        {
                            if (connection.State != ConnectionState.Open)
                                connection.Open();

                            using (SqlDataAdapter localAdapter = new SqlDataAdapter(cmd))
                            {
                                dataTable = new DataTable();
                                localAdapter.Fill(dataTable);
                                rownum = dataTable.Rows.Count;

                                DataTable dtAPA = dataTable.Clone();
                                DataTable dtAutre = dataTable.Clone();
                                DataTable dtLivre = dataTable.Clone();
                                DataTable dtCloture = dataTable.Clone();

                                dtLivre.Columns.Add("Action", typeof(string));

                                foreach (DataRow row in dataTable.Rows)
                                {
                                    string doPiece = row["DO_Piece"]?.ToString() ?? "";
                                    int doCloture = row["DO_Cloture"] == DBNull.Value ? 0 : Convert.ToInt32(row["DO_Cloture"]);

                                    if (doPiece.Contains("APA") || doPiece.Contains("BC"))
                                        dtAPA.ImportRow(row);
                                    else if (doPiece.Contains("AFA"))
                                        dtAutre.ImportRow(row);
                                    else if (doPiece.Contains("ABR"))
                                    {
                                        if (doCloture == 0)
                                        {
                                            DataRow newRow = dtLivre.NewRow();

                                            // Copy only the original columns (exclude Action)
                                            foreach (DataColumn col in dataTable.Columns)
                                                newRow[col.ColumnName] = row[col.ColumnName];

                                            newRow["Action"] = "Clôturer";
                                            dtLivre.Rows.Add(newRow);
                                        }
                                        else
                                        {
                                            dtCloture.ImportRow(row);
                                        }
                                    }
                                }
                                string[] val = { "DO_Cloture", "CO_No", "DE_No", "DO_TIERS","DO_Statut", "DO_Expedit","DO_Coord01","DO_Ref", "DO_DateLivr","DO_Date", "DO_DateExpedition", "A_TYPE", "DO_CodeTaxe1", "DO_Taxe1", "DO_Type", "DO_Imprim", "DO_Reliquat"};

                                foreach (DataTable dt in new[] { dtAPA, dtAutre, dtLivre, dtCloture })
                                {
                                    foreach (string valeur in val)
                                    {
                                        if (dt.Columns.Contains(valeur))
                                            dt.Columns[valeur].ColumnMapping = MappingType.Hidden;
                                    }
                                }

                                foreach (DataTable dt in new[] { dtAPA, dtAutre, dtLivre, dtCloture })
                                {
                                    if (dt.Columns.Contains("DO_Piece"))
                                        dt.Columns["DO_Piece"].Caption = "Numéro pièce";
                                    if (dt.Columns.Contains("DO_TotalTTC"))
                                        dt.Columns["DO_TotalTTC"].Caption = "Montant TTC";
                                    if (dt.Columns.Contains("CT_Intitule"))
                                        dt.Columns["CT_Intitule"].Caption = "Fournisseur";
                                }

                                // gc → APA
                                BindingSource bsAPA = new BindingSource();
                                bsAPA.DataSource = dtAPA;
                                gc.DataSource = bsAPA;  // toujours affecté, même vide

                                // gc1 → AFA
                                BindingSource bsAFA = new BindingSource();
                                bsAFA.DataSource = dtAutre;
                                gc1.DataSource = bsAFA;

                                // gc2 → ABR non clôturé
                                BindingSource bsLivre = new BindingSource();
                                bsLivre.DataSource = dtLivre;
                                gc2.DataSource = bsLivre;

                                // gc3 → ABR clôturé
                                BindingSource bsCloture = new BindingSource();
                                bsCloture.DataSource = dtCloture;
                                gc3.DataSource = bsCloture;

                                // Mettre à jour le bs original avec toutes les données
                                bs.DataSource = rownum > 0 ? (object)dataTable : null;
                            }
                        }
                    }
                    catch (SqlException ex) { MessageBox.Show("Erreur SQL : " + ex.Message); }
                    catch (Exception ex) { MessageBox.Show("Erreur : " + ex.Message); }
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static List<F_DEPOT> GetAllDepots()
        {
            List<F_DEPOT> depots = new List<F_DEPOT>();
            string query = "SELECT DE_No, DE_Intitule FROM F_DEPOT";


            using (SqlConnection conn = new SqlConnection(connectionStrings))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                depots.Add(new F_DEPOT
                                {
                                    DE_No = (int)reader["DE_No"],
                                    DE_Intitule = reader["DE_Intitule"].ToString()
                                });
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return depots;
        }
        public static List<P_UNITE> GetAllUnites()
        {
            List<P_UNITE> depots = new List<P_UNITE>();
            string query = "SELECT cbIndice, U_Intitule FROM P_UNITE WHERE U_Intitule <> ''";


            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                depots.Add(new P_UNITE
                                {
                                    cbIndice = (short)reader["cbIndice"],
                                    U_Intitule = reader["U_Intitule"].ToString()
                                });
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return depots;
        }

        public static List<P_DEVISE> GetAllDevise()
        {
            List<P_DEVISE> devises = new List<P_DEVISE>();
            string query = "SELECT cbMarq, D_Intitule FROM P_DEVISE WHERE D_Intitule <> ''";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                devises.Add(new P_DEVISE
                                {
                                    cbMarq = reader["cbMarq"] != DBNull.Value ? Convert.ToInt32(reader["cbMarq"]) : 0,
                                    D_Intitule = reader["D_Intitule"]?.ToString() ?? ""
                                });
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return devises;
        }

        public static List<P_EXPEDITION> GetAllExpedition()
        {
            List<P_EXPEDITION> expeditions = new List<P_EXPEDITION>();
            string query = "SELECT cbMarq, E_Intitule FROM P_EXPEDITION WHERE E_Intitule <> ''";


            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                expeditions.Add(new P_EXPEDITION
                                {
                                    cbMarq = reader["cbMarq"] != DBNull.Value ? Convert.ToInt32(reader["cbMarq"]) : 0,
                                    E_Intitule = reader["E_Intitule"].ToString()
                                });
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return expeditions;
        }
        public static void FiltrerFournisseurs(GridControl gc, bool actif, bool sommeil, BindingSource bsFrns)
        {
            try
            {
                DataTable dataTableFrns = new DataTable();

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "";
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = connection;

                    if (actif && !sommeil)
                    {
                        // Actif uniquement (CT_Sommeil = 0)
                        query = $"SELECT * FROM dbo.F_COMPTET WHERE CT_Sommeil = 0 AND CT_Type = 1 ORDER BY CT_Num";
                    }
                    else if (!actif && sommeil)
                    {
                        // En sommeil uniquement (CT_Sommeil = 1)
                        query = $"SELECT * FROM dbo.F_COMPTET WHERE CT_Sommeil = 1 AND CT_Type = 1 ORDER BY CT_Num";
                    }
                    else if (actif && sommeil)
                    {
                        // Les deux cochés, donc sans filtre
                        query = $"SELECT * FROM dbo.F_COMPTET WHERE CT_Type = 1 ORDER BY CT_Num";
                    }
                    else
                    {
                        // Aucun coché => vider la grille
                        gc.DataSource = null;
                        return;
                    }

                    cmd.CommandText = query;
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    connection.Open();
                    adapter.Fill(dataTableFrns);
                    rownum = dataTableFrns.Rows.Count;
                    bsFrns.DataSource = dataTableFrns;
                    gc.DataSource = bsFrns;

                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static List<F_COMPTEG> GetAllCompteG()
        {
            List<F_COMPTEG> comptetG = new List<F_COMPTEG>();
            string query = "SELECT distinct CG_Num, CG_Intitule FROM F_COMPTEG WHERE CG_Tiers = 1 AND CG_Sommeil = 0";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                comptetG.Add(new F_COMPTEG
                                {
                                    CG_Num = reader["CG_Num"].ToString(),
                                    CG_Intitule = reader["CG_Intitule"].ToString()
                                });
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return comptetG;
        }
    }
}
