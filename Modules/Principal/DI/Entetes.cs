using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Models;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
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
        private static SqlDataAdapter dataAdapter;

        public static void AfficherEntetes(GridControl gc, int achattype, BindingSource bs)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    string query = "";/// $"SELECT * FROM dbo.F_DOCENTETE WHERE DO_Domaine = 1 AND DO_Type = @dotype";
                    if(achattype != 200)
                    {
                        query =  $"SELECT * FROM dbo.ACHAT_ENTETE WHERE DO_Type = @dotype ORDER BY DO_Date DESC";
                    }
                    else
                    {
                        query = $"SELECT * FROM dbo.ACHAT_ENTETE ORDER BY DO_Date DESC";
                    }
                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@dotype", achattype);
                        dataAdapter = new SqlDataAdapter(cmd); // On passe `cmd`, pas `query`
                        SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                        dataTable = new DataTable();

                        connection.Open();
                        dataAdapter.Fill(dataTable);
                        rownum = dataTable.Rows.Count;
                        if(rownum == 0)
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
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static List<F_DEPOT> GetAllDepots()
        {
            List<F_DEPOT> depots = new List<F_DEPOT>();
            string query = "SELECT DE_No, DE_Intitule FROM F_DEPOT";


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
