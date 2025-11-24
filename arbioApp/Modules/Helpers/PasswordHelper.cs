using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;


namespace arbioApp
{
    public static class PasswordHelper
    {
        public static string HashPassword(string password)
        {
            if (string.IsNullOrEmpty(password))
                throw new ArgumentNullException(nameof(password), "Le mot de passe ne peut pas être vide.");

            using (SHA256 sha256 = SHA256.Create())
            {
                // Convertir le mot de passe en tableau de bytes
                byte[] passwordBytes = Encoding.UTF8.GetBytes(password);

                // Calculer le hash
                byte[] hashBytes = sha256.ComputeHash(passwordBytes);

                // Convertir le hash en chaîne de caractères hexadécimale
                StringBuilder hashString = new StringBuilder();
                foreach (byte b in hashBytes)
                {
                    hashString.Append(b.ToString("x2")); // Convertir chaque byte en une valeur hexadécimale
                }

                return hashString.ToString();
            }
        }
    }
}
