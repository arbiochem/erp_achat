using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal
{
    public class BrowsSites
    {

        public static DataTable ExecuteQueryOnMultipleServers(string mainConnectionString, string query)
        {
            DataTable combinedResults = new DataTable();

            using (SqlConnection mainConnection = new SqlConnection(mainConnectionString))
            {
                mainConnection.Open();

                string serverQuery = "SELECT ADDRESS_IP, BDD FROM dbo.T_SERVER_DEPOT WHERE actif = 1";
                SqlCommand serverCommand = new SqlCommand(serverQuery, mainConnection);

                using (SqlDataReader reader = serverCommand.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string serverAddress = reader["ADDRESS_IP"].ToString();
                        string databaseName = reader["BDD"].ToString();

                        string targetConnectionString = $"Data Source={serverAddress};Initial Catalog={databaseName};User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=120;";

                        using (SqlConnection targetConnection = new SqlConnection(targetConnectionString))
                        {
                            try
                            {
                                targetConnection.Open();

                                SqlDataAdapter adapter = new SqlDataAdapter(query, targetConnection);
                                DataTable serverResult = new DataTable();
                                adapter.Fill(serverResult);

                                combinedResults.Merge(serverResult);

                            }
                            catch (Exception ex)
                            {
                                MethodBase m = MethodBase.GetCurrentMethod();
                                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }

            return combinedResults;
        }
    }
}
