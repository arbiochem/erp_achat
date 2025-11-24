using arbioApp.Modules.Principal.DI._2_Documents;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Utils.Connection
{
    public static class Db
    {
        public static async Task<(SqlConnection connection, string connectionString)> ConnectToServer(string serverName, string userName, string password, int timeoutInSeconds)
        {
            string connectionString = $"Data Source={serverName};User ID={userName};Password={password};TrustServerCertificate=True;";
            if (userName == "" && password == "")
            {
                connectionString = $"Data Source={serverName};Integrated Security=True;TrustServerCertificate=False;";
            }
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    using (CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutInSeconds)))
                    {
                        await connection.OpenAsync(cts.Token);
                        // Close the connection when done
                        
                        connection.Close();
                        MessageBox.Show("Connecté avec succès.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                        return (connection, connectionString);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException($"Connection attempt timed out after {timeoutInSeconds} seconds.");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error: {ex.Message}");
            }
        }


        private static string DbPrincipale => ucDocuments.dbNamePrincipale;
        private static string ServerIpPrincipale => ucDocuments.serverIpPrincipale;
        public static string GetConnectionString()
        {
            if (string.IsNullOrEmpty(DbPrincipale) || string.IsNullOrEmpty(ServerIpPrincipale))
                throw new InvalidOperationException("La base de données ou le serveur n'ont pas encore été sélectionnés.");

            string connectionString = $"Server={ServerIpPrincipale};Database={DbPrincipale};" +
                                      $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                      $"Connection Timeout=240;";

            return connectionString;
        }
    }
}
