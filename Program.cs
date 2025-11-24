using DevExpress.LookAndFeel;
using DevExpress.Skins;
using DevExpress.UserSkins;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace arbioApp
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
            //string appName = Process.GetCurrentProcess().ProcessName;

            //var runningProcesses = Process.GetProcessesByName(appName).Where(p => p.Id != Process.GetCurrentProcess().Id);

            //// Fermer les processus existants
            //foreach (var process in runningProcesses)
            //{
            //    try
            //    {
            //        process.Kill(); // Fermer le processus
            //        process.WaitForExit(); // Attendre que le processus soit complètement fermé
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show($"Impossible de fermer l'application en cours : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        return; // Quitter si on ne peut pas fermer l'instance
            //    }
            //}
            Application.Run(new FrmConnex());
        }
    }
}
