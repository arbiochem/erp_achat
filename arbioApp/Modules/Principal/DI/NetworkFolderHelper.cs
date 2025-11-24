using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace arbioApp.Modules.Principal.DI
{
    

    public static class NetworkFolderHelper
    {
        /// <summary>
        /// Vérifie que le dossier distant est accessible, le crée si nécessaire,
        /// puis copie tous les fichiers depuis un dossier source.
        /// </summary>
        /// <param name="sourceFolder">Dossier local contenant les fichiers à transférer.</param>
        /// <param name="destinationFolder">Dossier distant (UNC) ex: \\Serveur\Partage\Dossier</param>
        /// <param name="overwrite">Remplacer les fichiers existants ?</param>
        /// <returns>true si succès, false sinon</returns>
        public static bool TransferFiles(string sourceFolder, string destinationFolder, bool overwrite = true)
        {
            try
            {
                // Vérifie si la source existe
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine("❌ Dossier source introuvable : " + sourceFolder);
                    return false;
                }

                // Crée le dossier distant si besoin
                if (!Directory.Exists(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                    Console.WriteLine("📁 Dossier distant créé : " + destinationFolder);
                }

                // Copie les fichiers
                foreach (var file in Directory.GetFiles(sourceFolder))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(destinationFolder, fileName);

                    File.Copy(file, destFile, overwrite);
                    Console.WriteLine($"✅ {fileName} copié vers {destinationFolder}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erreur lors du transfert : " + ex.Message);
                return false;
            }
        }
    }

}
