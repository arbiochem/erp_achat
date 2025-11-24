using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.CodeParser;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraLayout.Customization;
using DevExpress.XtraSplashScreen;
using Microsoft.Win32;
using System.Data.Sql;
using System.Data.SqlClient;
using System.IO;
using DevExpress.XtraTreeList.Nodes;
using System.Reflection;

namespace arbioApp
{
    public static class Addserver
    {
        public static void LoadSqlServers(DevExpress.XtraEditors.ComboBoxEdit cboServers)
        {

            string filePath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "serveurs.txt");
            string[] lines = File.ReadAllLines(filePath);

            // Ajouter chaque ligne au ComboBox
            foreach (string line in lines)
            {
                cboServers.Properties.Items.Add(line);
            }
        }
        public static void LoadServerConnex(DevExpress.XtraEditors.ComboBoxEdit cboServers)
        {

            string filePath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "serveurs_connex.txt");
            string[] lines = File.ReadAllLines(filePath);

            // Ajouter chaque ligne au ComboBox
            foreach (string line in lines)
            {
                cboServers.Properties.Items.Add(line);
            }
        }

        public static DataTable GetDatabasesForServer(string serverName)
        {
            DataTable dataTable = new DataTable();

            try
            {
                string connectionString = $"Data Source={serverName};Initial Catalog=master;User ID=Dev;Password=1234;TrustServerCertificate=True";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open(); // S'assurer que la connexion fonctionne
                    string query = "SELECT name AS DatabaseName FROM sys.databases WHERE database_id > 4"; // Ignorer les bases système
                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    adapter.Fill(dataTable);
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }

            return dataTable;
        }


        //public static void LoadTreeList(DevExpress.XtraTreeList.TreeList treeList)
        //{
        //    try
        //    {
        //        treeList.Nodes.Clear(); // Réinitialiser le TreeList

        //        // Charger la liste des serveurs
        //        string filePath = System.IO.Path.Combine(Application.StartupPath, "serveurs.txt");
        //        if (!File.Exists(filePath))
        //        {
        //            MessageBox.Show("Le fichier serveurs.txt est introuvable.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return;
        //        }

        //        string[] servers = File.ReadAllLines(filePath);
        //        if (servers.Length == 0)
        //        {
        //            MessageBox.Show("Aucun serveur trouvé dans le fichier serveurs.txt.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            return;
        //        }

        //        // Ajouter des colonnes si elles n'existent pas déjà
        //        if (treeList.Columns.Count == 0)
        //        {
        //            treeList.Columns.Add();
        //            treeList.Columns[0].Caption = "Nom";
        //            treeList.Columns[0].Visible = true;
        //        }

        //        foreach (string server in servers)
        //        {
        //            // Ajouter le serveur comme nœud parent
        //            TreeListNode serverNode = treeList.AppendNode(new object[] { server }, null);
        //            serverNode.Tag = server; // Enregistrer le nom du serveur dans le Tag

        //            // Charger les bases de données pour ce serveur
        //            DataTable databases = GetDatabasesForServer(server);
        //            if (databases.Rows.Count > 0)
        //            {
        //                foreach (DataRow row in databases.Rows)
        //                {
        //                    string dbName = row["DatabaseName"].ToString();

        //                    // Ajouter la base de données comme nœud enfant
        //                    TreeListNode databaseNode = treeList.AppendNode(new object[] { dbName }, serverNode);
        //                    databaseNode.Tag = dbName; // Enregistrer le nom de la base dans le Tag
        //                }
        //            }
        //            else
        //            {
        //                treeList.AppendNode(new object[] { "(Aucune base trouvée)" }, serverNode);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Erreur lors du chargement du TreeList : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        throw;
        //    }
        //}
        public static void LoadTreeList(string serverName, DevExpress.XtraTreeList.TreeList treeList)
        {
            // Connexion à SQL Server
            string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";
            string query = "SELECT Id, ParentID, Site, connex FROM T_Server";

            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(command);
                        adapter.Fill(dataTable);
                    }
                }

                // Configuration du TreeList
                treeList.BeginUpdate();
                treeList.DataSource = dataTable;

                //// Définir les champs hiérarchiques
                treeList.KeyFieldName = "Id";
                treeList.ParentFieldName = "ParentID";

                //// Ajouter les colonnes explicitement si nécessaire
                //treeList.Columns.Clear();
                ////treeList.Columns.AddVisible("Site", "Site"); // Colonne affichée
                ////treeList.Columns.AddVisible("connex", "Connexion"); // Colonne additionnelle
                treeList.Columns["connex"].VisibleIndex = -1;// Masquer la colonne 'connex'

                treeList.OptionsBehavior.Editable = false; // Rendre le TreeList non éditable
                treeList.OptionsView.ShowIndicator = false; // Option pour masquer l'indicateur de ligne
                treeList.EndUpdate();


            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        public static void LoadComboSites(DevExpress.XtraEditors.ComboBoxEdit cboSite)
        {
            string filePath = @"cboSites.txt";

            try
            {
                // Vérifier si le fichier existe
                if (File.Exists(filePath))
                {
                    // Lire toutes les lignes du fichier
                    string[] lines = File.ReadAllLines(filePath);

                    // Parcourir chaque ligne et l'ajouter au ComboBoxEdit
                    foreach (string line in lines)
                    {
                        if (line.StartsWith("\t") || line.StartsWith("    ")) // Identifier les lignes enfants
                        {
                            cboSite.Properties.Items.Add("   " + line.Trim()); // Indenter visuellement
                        }
                        else
                        {
                            cboSite.Properties.Items.Add(line.Trim()); // Ajouter normalement les parents
                        }
                    }

                    // Optionnel : Sélectionner le premier élément par défaut
                    if (cboSite.Properties.Items.Count > 0)
                    {
                        cboSite.SelectedIndex = 0;
                    }
                }
                else
                {
                    MessageBox.Show("Le fichier texte spécifié est introuvable.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
        public static void LoadTreeListLookUpEdit(string serverName, DevExpress.XtraEditors.TreeListLookUpEdit treeListLookUpEdit)
        {
            // Connexion à SQL Server
            string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";
            string query = "SELECT Id, ParentID, Site, connex FROM T_Server";

            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(command);
                        adapter.Fill(dataTable);
                    }
                }

                // Mise à jour du TreeListLookUpEdit
                treeListLookUpEdit.Properties.TreeList.BeginUpdate();

                // Ajouter les colonnes explicitement
                //treeListLookUpEdit.Properties.TreeList.Columns.Clear();

                treeListLookUpEdit.Properties.TreeList.KeyFieldName = "Id";
                treeListLookUpEdit.Properties.TreeList.ParentFieldName = "ParentID";
                treeListLookUpEdit.Properties.DisplayMember = "Site";
                treeListLookUpEdit.Properties.ValueMember = "Id";
                treeListLookUpEdit.Properties.TreeList.DataSource = dataTable;
                treeListLookUpEdit.Properties.TreeList.Columns["connex"].VisibleIndex = -1;
                treeListLookUpEdit.Properties.TreeList.EndUpdate();
                // Désactiver la sélection des nœuds parent
                treeListLookUpEdit.Properties.TreeList.BeforeFocusNode += (sender, e) =>
                {
                    var treeList = sender as DevExpress.XtraTreeList.TreeList;
                    if (treeList != null)
                    {
                        // Vérifie si le nœud sélectionné a des enfants
                        if (treeList.GetNodeList().Any(node => node.ParentNode?.Id == e.Node.Id))
                        {
                            e.CanFocus = false; // Empêche la sélection du nœud
                        }
                    }
                };

            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        public static void LoadLookUpEditSite(string serverName, LookUpEdit lookUpEdit)
        {
            string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";
            string query = "SELECT * FROM T_SITE_CATTARIF";  

            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();  
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(command);
                        adapter.Fill(dataTable);
                    }
                }

                if (dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("Aucune donnée trouvée dans T_SITE_CATTARIF.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                lookUpEdit.Properties.BeginUpdate();

                lookUpEdit.Properties.Columns.Clear(); 
                lookUpEdit.Properties.DataSource = dataTable;
                lookUpEdit.Properties.DisplayMember = "SiteName"; 
                lookUpEdit.Properties.ValueMember = "SiteName";
                lookUpEdit.Properties.Columns.Add(new LookUpColumnInfo("SiteName", "Nom du Site"));
                lookUpEdit.Properties.Columns.Add(new LookUpColumnInfo("Address_IP", "Adresse IP"));

                lookUpEdit.Properties.EndUpdate(); 
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static void LoadTreeListDepot(string serverName, DevExpress.XtraTreeList.TreeList treeList)
        {
            // Connexion à SQL Server
            string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";
            string query = "SELECT * FROM T_SERVER_DEPOT";

            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(command);
                        adapter.Fill(dataTable);
                    }
                }

                // Configuration du TreeList
                treeList.BeginUpdate();
                treeList.DataSource = dataTable;

                treeList.Columns["ADDRESS_IP"].VisibleIndex = -1;// Masquer la colonne 
                treeList.Columns["LOCATION"].VisibleIndex = -1;// Masquer la colonne 

                treeList.OptionsBehavior.Editable = false; // Rendre le TreeList non éditable
                treeList.OptionsView.ShowIndicator = false; // Option pour masquer l'indicateur de ligne
                treeList.EndUpdate();


            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }


}


