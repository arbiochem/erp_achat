using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace arbioApp.Utils
{
    public static class Cmbx
    {
        public static async Task<string[]> Populate(SqlConnection connection, string connectionString)
        {
            using (connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("SELECT name FROM sys.databases", connection))
                {
                    using (SqlDataReader reader = await command.ExecuteReaderAsync())
                    {
                        var databases = new System.Collections.Generic.List<string>();
                        while (await reader.ReadAsync())
                        {
                            string databaseName = reader["name"].ToString();
                            databases.Add(databaseName);
                        }
                        connection.Close();
                        return databases.ToArray();
                    }
                }
                
            }
            
        }
        public static int GetValueMember(ComboBox cmbx, string valueMember, string displayMember)
        {
            int IdRetour = 0;
            var selectedItem = cmbx.SelectedItem;
            if (selectedItem != null)
            {
                // Split the selected text to extract RoleId
                var selectedValue = cmbx.SelectedItem.ToString();
                int startIndex = selectedValue.IndexOf(""+valueMember+" = ") + (valueMember+" = ").Length;
                int endIndex = selectedValue.IndexOf(", "+displayMember+"");
                string roleIdString = selectedValue.Substring(startIndex, endIndex - startIndex).Trim();
                //int.TryParse(roleIdString, out IdRetour);
                IdRetour = int.Parse(roleIdString);
            }
            return IdRetour;
        }
    }
}
