
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Sql;
using System.Data.SqlClient;
namespace arbioApp.Modules.Helpers
{
    public class DateHelper
    {
        public static void ApposerLastDate(string serverName, string Username)
        {
            string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string queryUpdateLastDate = "UPDATE T_UserRole SET LastDateLogon = @LastDateLogon WHERE Username = @Username";
                using (SqlCommand command = new SqlCommand(queryUpdateLastDate, connection))
                {
                    command.Parameters.AddWithValue("@LastDateLogon", DateTime.Now); // Hash du nouveau mot de passe
                    command.Parameters.AddWithValue("@Username", Username);

                    command.ExecuteNonQuery();

                }
            }
        }
        public static void ApposerCreateDate(string serverName, string Username)
        {
            string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string queryUpdateLastDate = "UPDATE T_UserRole SET CreateDate = @CreateDate WHERE Username = @Username";
                using (SqlCommand command = new SqlCommand(queryUpdateLastDate, connection))
                {
                    command.Parameters.AddWithValue("@CreateDate", DateTime.Now); // Hash du nouveau mot de passe
                    command.Parameters.AddWithValue("@Username", Username);

                    command.ExecuteNonQuery();

                }
            }
        }
    }
}
