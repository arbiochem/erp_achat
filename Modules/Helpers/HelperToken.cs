using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Helpers
{
    public static class HelperToken
    {
        public static void LoadEmailsFromDatabase(TokenEdit tokenEdit)
        {
            string serverName = File.ReadLines("servermail.txt").FirstOrDefault();

            if (serverName != null)
            {
                Console.WriteLine("Première ligne : " + serverName);
            }
            else
            {
                Console.WriteLine("Le fichier est vide.");
            }

            string connectionString = $"Data Source={serverName};Initial Catalog=arbApp;User ID=Dev;Password=1234;TrustServerCertificate=True";
            string query = "SELECT Username FROM T_UserRole WHERE Username IS NOT NULL";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();


                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            tokenEdit.Properties.Tokens.Clear();
                            while (reader.Read())
                            {
                                string email = reader["Username"].ToString();
                                tokenEdit.Properties.Tokens.Add(new TokenEditToken(email, email));
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des emails : {ex.Message} " + Environment.NewLine + "Veuillez réessayer plus tard !");
                //Environment.Exit(0);
            }
        }
        public static void TokenEdit_ValidateToken(object sender, TokenEditValidateTokenEventArgs e)
        {
            if (IsValidEmail(e.Value.ToString()))
            {
                e.IsValid = true;
                e.Description = e.Value.ToString();
            }
            else
            {
                e.IsValid = false;
            }
        }
        public static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
}
