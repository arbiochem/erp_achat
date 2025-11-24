using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraReports.Design;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.Dashboard
{
    public class clsMultiDataTable
    {
        public static DataTable ExecuteQueryOnMultipleServers(string query)
        {
            DataTable combinedResults = new DataTable();

            string mainConnectionString = $"Data Source={FrmMdiParent.DataSourceNameValueParent};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=120;";

            using (SqlConnection mainConnection = new SqlConnection(mainConnectionString))
            {
                mainConnection.Open();

                string serverQuery = "SELECT ADDRESS_IP, BDD FROM dbo.T_SERVER_REPORT WHERE actif = 1";
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
                                MessageBox.Show($"[{DateTime.Now}]: {serverAddress}: {ex.Message}");
                            }
                        }
                    }
                }
            }

            return combinedResults;
        }
    }
}
