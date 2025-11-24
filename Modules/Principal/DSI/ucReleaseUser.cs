using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraSpreadsheet;
using DataTable = System.Data.DataTable;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using System.Diagnostics;
using DevExpress.XtraBars.Customization;
using DevExpress.XtraPivotGrid;
using arbioApp.Modules.Helpers;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;
using System.Drawing;
using arbioApp.Modules.Principal.GESTION_COMMERCIALE;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Grid;
using System.Reflection;
using arbioApp.Modules.Principal.Dashboard.RECOUVREMENT;

namespace arbioApp.Modules.Principal.DSI
{
    public partial class ucReleaseUser : DevExpress.XtraEditors.XtraUserControl
    {
        private static ucReleaseUser _instance;
        public static ucReleaseUser Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucReleaseUser();
                return _instance;
            }
        }

        public ucReleaseUser()
        {
            InitializeComponent();
            
        }
        private void LoadcbUserSession()
        {
            try
            {
                gridControl4.DataSource = null;
                gridView4.Columns.Clear();

                string leserveur = FrmMdiParent.DataSourceNameValueParent;
                string connectionString = $"Data Source={leserveur};Initial Catalog=ARBIOCHEM;User ID=Dev;Password=1234;TrustServerCertificate=True";

                string query = "SELECT cbSession, CB_Type, cbUserName FROM cbUserSession WHERE cbUserName is not null";
                DataTable dataTable = new DataTable();
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    adapter.Fill(dataTable);
                }

                gridControl4.DataSource = dataTable;

                // ✅ Force la création des colonnes avant de les utiliser
                gridControl4.ForceInitialize();

                gridView4.BeginSort();
                try
                {
                    gridView4.ClearGrouping();
                    gridView4.GroupCount = 1;
                    gridView4.Columns["CB_Type"].GroupIndex = 0;

                }
                finally
                {
                    gridView4.EndSort();
                }
                gridView4.ExpandAllGroups();
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ucReleaseUser_Load(object sender, EventArgs e)
        {
            Addserver.LoadSqlServers(CboReleaseUser);
            string leserveur = FrmMdiParent.DataSourceNameValueParent;
            string UserRole = FrmMdiParent.UserRole;

            Addserver.LoadTreeListLookUpEdit(leserveur, treeListLookUpEdit1);
            treeListLookUpEdit1.Properties.AutoExpandAllNodes = true;

            LoadcbUserSession();
            contextMenuStrip1.Items.Add("Déconnecter cette session !", null, SupprimerLigne_Click);

        }
        public DataTable GetDatabases(string connectionString)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT name AS DatabaseName FROM sys.databases WHERE database_id > 4 ORDER BY DatabaseName";
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                adapter.Fill(dataTable); // Si une erreur se produit ici, elle remontera
            }

            return dataTable;
        }


        string serverName;
        private void HyperlinkEdit_OpenLink(object sender, DevExpress.XtraEditors.Controls.OpenLinkEventArgs e)
        {
            // Récupérer le nom de la base de données
            //GridView view = gridControl1.FocusedView as GridView;
            int rowHandle = gridView1.FocusedRowHandle;
            string databaseName = gridView1.GetRowCellValue(rowHandle, "DatabaseName").ToString();

            ReleaseUser(serverName, databaseName);

        }
        private void ReleaseUser(string serverName, string databaseName)
        {
            string connectionStringTemplate = $"Data Source={serverName};Initial Catalog={databaseName};User ID=Dev;Password=1234;TrustServerCertificate=True";

            // Construire la chaîne de connexion avec la base de données sélectionnée
            string connectionString = string.Format(connectionStringTemplate, databaseName);

            string script = @"
                                DBCC CBSQLXP (FREE);
                                DELETE FROM cbMessage;
                                DELETE FROM cbNotification;
                                DELETE FROM cbRegFile;
                                DELETE FROM cbRegMessage;
                                DELETE FROM cbRegUser;
                                DELETE FROM cbUserSession;";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(script, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Script exécuté avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur est survenue : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                serverName = CboReleaseUser.Text;
                string connectionString = $"Data Source={serverName};Initial Catalog=master;User ID=Dev;Password=1234;TrustServerCertificate=True";
                //string connectionString = $"Data Source={serverName};Initial Catalog=master;User ID=SRV-ANTANIMORA\\Administrateur;";
                DataTable databases = GetDatabases(connectionString);

                // Configurer le GridControl
                gridControl1.DataSource = databases;

                // Ajouter un RepositoryItemHyperLinkEdit pour rendre les cellules cliquables
                RepositoryItemHyperLinkEdit hyperlinkEdit = new RepositoryItemHyperLinkEdit();
                gridView1.Columns["DatabaseName"].ColumnEdit = hyperlinkEdit;

                // Ajouter un événement de clic
                hyperlinkEdit.OpenLink += HyperlinkEdit_OpenLink;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur est survenue : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {
                splashScreenManager1.ShowWaitForm();


                serverName = CboReleaseUser.Text;
                string connectionString = $"Data Source={serverName};Initial Catalog=master;User ID=Dev;Password=1234;TrustServerCertificate=True";
                //string connectionString = $"Data Source={serverName};Initial Catalog=master;User ID=SRV-ANTANIMORA\\Administrateur;";
                DataTable databases = GetDatabases(connectionString);

                // Configurer le GridControl
                gridControl1.DataSource = databases;

                // Ajouter un RepositoryItemHyperLinkEdit pour rendre les cellules cliquables
                RepositoryItemHyperLinkEdit hyperlinkEdit = new RepositoryItemHyperLinkEdit();
                gridView1.Columns["DatabaseName"].ColumnEdit = hyperlinkEdit;

                // Ajouter un événement de clic
                hyperlinkEdit.OpenLink += HyperlinkEdit_OpenLink;
                splashScreenManager1.CloseWaitForm();
            }
            catch (Exception ex)
            {
                splashScreenManager1.CloseWaitForm();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void simpleButtonprevue_Click(object sender, EventArgs e)
        {
            gridControl1.ShowPrintPreview();
        }

        private void treeListDataSource_BeforeCheckNode(object sender, DevExpress.XtraTreeList.CheckNodeEventArgs e)
        {
            TreeListNode node = e.Node;
            if (node.Checked)
            {
                node.UncheckAll();
            }
            else
            {
                node.CheckAll();
            }
            while (node.ParentNode != null)
            {
                node = node.ParentNode;
                bool oneOfChildIsChecked = OneOfChildsIsChecked(node);
                if (oneOfChildIsChecked)
                {
                    node.CheckState = CheckState.Checked;
                }
                else
                {
                    node.CheckState = CheckState.Unchecked;
                }
            }
        }
        private bool OneOfChildsIsChecked(TreeListNode node)
        {
            bool result = false;
            foreach (TreeListNode item in node.Nodes)
            {
                if (item.CheckState == CheckState.Checked)
                {
                    result = true;
                }
            }
            return result;
        }

        private void RecupererAncienPrix()
        {

        }


        private void treeListDataSource_CustomNodeCellEdit(object sender, DevExpress.XtraTreeList.GetCustomNodeCellEditEventArgs e)
        { }

        private void simpleButton1_Click_1(object sender, EventArgs e)
        {

        }

        public DataTable ExtractSpreadsheetData(SpreadsheetControl spreadsheet)
        {
            var workbook = spreadsheet.Document;
            var worksheet = workbook.Worksheets.ActiveWorksheet;
            var dataTable = new DataTable();

            // Ajouter les colonnes du DataTable (en fonction des colonnes du Spreadsheet)
            dataTable.Columns.Add("AR_Ref", typeof(string));
            dataTable.Columns.Add("AR_Design", typeof(string));
            dataTable.Columns.Add("AR_PrixVen", typeof(decimal));
            dataTable.Columns.Add("CatTarif", typeof(string));
            dataTable.Columns.Add("NouveauPrix", typeof(decimal));

            // Parcourir les lignes pour récupérer les données
            int startRow = 1; // Supposons que la première ligne contient les en-têtes
            int lastRow = worksheet.Rows.LastUsedIndex;

            for (int rowIndex = startRow + 1; rowIndex <= lastRow; rowIndex++) // Ignorer les en-têtes
            {
                var row = worksheet.Rows[rowIndex];
                var newRow = dataTable.NewRow();

                newRow["AR_Ref"] = row[0].Value.TextValue;
                newRow["AR_Design"] = row[1].Value.TextValue;
                newRow["AR_PrixVen"] = row[2].Value.NumericValue;
                newRow["CatTarif"] = row[3].Value.TextValue;
                newRow["NouveauPrix"] = row[4].Value.NumericValue;

                dataTable.Rows.Add(newRow);
            }

            return dataTable;
        }
        public void UpdateDatabaseFromSpreadsheet(SpreadsheetControl spreadsheet, string connectionString)
        {
            // Extraire les données du SpreadsheetControl
            DataTable dataTable = ExtractSpreadsheetData(spreadsheet);

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Parcourir chaque ligne du DataTable
                foreach (DataRow row in dataTable.Rows)
                {
                    try
                    {
                        string arRef = row["AR_Ref"].ToString();
                        string catTarif = row["CatTarif"].ToString();
                        decimal nouveauPrix = Convert.ToDecimal(row["NouveauPrix"]);

                        // Créer la commande SQL
                        string updateQuery = @"UPDATE f
                                    SET f.AC_PrixVen = @NouveauPrix
                                    FROM F_ARTCLIENT f
                                    LEFT JOIN P_CATTARIF catTarif ON f.AC_Categorie = catTarif.cbIndice
                                    WHERE f.AR_Ref = @AR_Ref AND catTarif.CT_Intitule = @CatTarif;";

                        using (var command = new SqlCommand(updateQuery, connection))
                        {
                            // Ajouter les paramètres
                            command.Parameters.AddWithValue("@NouveauPrix", nouveauPrix);
                            command.Parameters.AddWithValue("@AR_Ref", arRef);
                            command.Parameters.AddWithValue("@CatTarif", catTarif);
                            int rowsAffected = command.ExecuteNonQuery();
                            MessageBox.Show($"Mise à jour effectuée. Lignes affectées : {rowsAffected}");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erreur lors de la mise à jour : {ex.Message}");
                    }
                }

            }
        }

        private DataSet dataSetCrud;
        private SqlConnection connection;
        private SqlDataAdapter adapterCrud;
        //private DataTable dataTable;


        public static string selectedSite;
        public static string selectedConnex;

        private void LoadData()
        {
            try
            {
                gridControl3.DataSource = null;
                gridView3.Columns.Clear();
                if (treeListLookUpEdit1 != null && treeListLookUpEdit1.Properties.TreeList.FocusedNode != null)
                {
                    // Récupérer la ligne sélectionnée (DataRowView)
                    var selectedRow = treeListLookUpEdit1.Properties.TreeList.GetDataRecordByNode(
                        treeListLookUpEdit1.Properties.TreeList.FocusedNode
                    ) as DataRowView;

                    if (selectedRow != null)
                    {
                        selectedSite = selectedRow["Site"].ToString();
                        selectedConnex = selectedRow["Connex"].ToString();

                        splashScreenManager1.ShowWaitForm();

                        try
                        {

                            string cs = $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True;";
                            using (SqlConnection conn = new SqlConnection(cs))
                            using (SqlCommand cmd = new SqlCommand("dbo.GetDepotAuthorizedPivot", conn))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;

                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                DataTable dt = new DataTable();
                                da.Fill(dt);

                                
                            
                            
                            GridHScrollHelper helper = new GridHScrollHelper(gridView3);
                            helper.EnableScrollByColumns();
                            gridControl3.DataSource = dt;
                            
                            gridView3.Columns["PROT_Guid"].VisibleIndex = -1;


                            gridView3.Columns["PROT_User"].Width = 200;
                            gridView3.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

                            RepositoryItemCheckEdit checkEdit = new RepositoryItemCheckEdit();
                            checkEdit.ValueChecked = 1;
                            checkEdit.ValueUnchecked = 0;
                            checkEdit.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;

                            gridControl3.RepositoryItems.Add(checkEdit);
                            for (int i = 2; i < gridView3.Columns.Count; i++)
                            {
                                gridView3.Columns[i].ColumnEdit = checkEdit;
                            }
                            checkEdit.CheckStateChanged += CheckEdit_CheckStateChanged;

                            }
                        }
                        finally
                        {
                            splashScreenManager1.CloseWaitForm();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SplashScreenManager.CloseForm(false);
            }
        }


        private void treeListLookUpEdit1_EditValueChanged(object sender, EventArgs e)
        {
            LoadData();

        }


        private void gridControl2_EmbeddedNavigator_ButtonClick(object sender, NavigatorButtonClickEventArgs e)
        {
            switch (e.Button.ButtonType)
            {
                case NavigatorButtonType.Append:
                    AddEntry();
                    e.Handled = true;
                    break;

                case NavigatorButtonType.Remove:
                    DeleteEntry();
                    e.Handled = true;
                    break;

                case NavigatorButtonType.EndEdit:
                    //UpdateEntry();
                    e.Handled = true;
                    break;

                default:
                    break;
            }
        }
        private void AddEntry()
        {
            // Récupérer les valeurs des colonnes directement à partir de GridView
            string deNo = gridView2.GetFocusedRowCellValue("DE_NO")?.ToString();
            string authorized = gridView2.GetFocusedRowCellValue("AUTHORIZED")?.ToString();
            string protGuid = gridView2.GetFocusedRowCellValue("PROT_Guid")?.ToString();

            string query = "INSERT INTO dbo.F_DEPOT_DEDIE (DE_NO, AUTHORIZED, PROT_Guid) VALUES (@DE_NO, @AUTHORIZED, @PROT_Guid)";
            ExecuteNonQuery(
                query,
                new SqlParameter("@DE_NO", deNo),
                new SqlParameter("@AUTHORIZED", authorized),
                new SqlParameter("@PROT_Guid", protGuid)
            );

            //LoadData(); // Rechargez les données après l'ajout
            //gridView2.AddNewRow();
        }
        private void DeleteEntry()
        {
            // Récupérer l'ID de la ligne sélectionnée
            int id = Convert.ToInt32(gridView2.GetFocusedRowCellValue("ID"));

            // Afficher un message de confirmation avant de supprimer
            DialogResult result = MessageBox.Show(
                "Êtes-vous sûr de vouloir supprimer cet enregistrement ?",
                "Confirmation de suppression",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            // Si l'utilisateur confirme, continuer la suppression
            if (result == DialogResult.Yes)
            {
                string connectionString =
                    $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=0;";
                string query = "DELETE FROM dbo.F_DEPOT_DEDIE WHERE ID = @ID";

                using (var connection = new SqlConnection(connectionString))
                using (var command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    command.Parameters.AddWithValue("@ID", id);
                    int rowsAffected = command.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Enregistrement supprimé avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Aucune suppression n'a été effectuée. Veuillez vérifier les données.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // Recharger les données après suppression
                //LoadData();
            }
            else
            {
                MessageBox.Show("Suppression annulée.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void UpdateEntry(int id, string deNo, string authorized, string protGuid)
        {
            if (id > 0)
            {
                string connectionString =
                    $"Data Source={selectedConnex};Initial Catalog={selectedSite};" +
                    $"User ID=Dev;Password=1234;TrustServerCertificate=True";
                string query = @"UPDATE dbo.F_DEPOT_DEDIE
                                SET
                                  DE_NO = @DE_NO,
                                  AUTHORIZED = @AUTHORIZED,
                                  PROT_Guid = @PROT_Guid
                                WHERE ID = @ID";
                using (var command = new SqlCommand(query, connection = new SqlConnection(connectionString)))
                {
                    connection.Open();
                    command.Parameters.AddWithValue("@DE_NO", deNo);
                    command.Parameters.AddWithValue("@AUTHORIZED", authorized);//authorized
                    command.Parameters.AddWithValue("@PROT_Guid", protGuid);
                    command.Parameters.AddWithValue("@ID", id);
                    int rowsAffected = command.ExecuteNonQuery();

                }

                //LoadData(); // Rechargez les données après la mise à jour
            }
            else
            {
                MessageBox.Show("Execute INSERT INTO F_DEPOT_DEDIE (DE_NO, AUTHORIZED, PROT_Guid) to F_DEPOT_DEDIE");
            }
        }
        private void ExecuteNonQuery(string query, params SqlParameter[] parameters)
        {
            try
            {
                // Remplacez "connectionString" par votre chaîne de connexion à SQL Server
                string connectionString =
                    $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";


                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand(query, connection);
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }

                    connection.Open(); // Ouvrir la connexion
                    command.ExecuteNonQuery(); // Exécuter la commande SQL
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            LoadData();
        }
        private DataTable GetDepots()
        {
            string connectionString =
                $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";

            string query = "SELECT DE_NO, DE_Intitule FROM dbo.F_DEPOT";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                return dataTable;
            }
        }
        private DataTable GetUsers()
        {
            string connectionString =
                $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";

            string query = "SELECT PROT_Guid, PROT_User FROM dbo.F_PROTECTIONCIAL";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                return dataTable;
            }
        }


        private void repositoryItemLookUpEdit3_EditValueChanged(object sender, EventArgs e)
        {
            object value = (sender as LookUpEdit).EditValue;
            int id = Convert.ToInt32(gridView2.GetFocusedRowCellValue("ID"));
            string deNo = gridView2.GetFocusedRowCellValue("DE_NO")?.ToString();
            string authorized = gridView2.GetFocusedRowCellValue("AUTHORIZED")?.ToString();
            //string protGuid = gridView2.GetFocusedRowCellValue("PROT_Guid")?.ToString();
            UpdateEntry(id, deNo, authorized, value.ToString());
        }

        private void repositoryItemLookUpEdit4_EditValueChanged(object sender, EventArgs e)
        {
            object value = (sender as LookUpEdit).EditValue;
            int id = Convert.ToInt32(gridView2.GetFocusedRowCellValue("ID"));
            //string deNo = gridView2.GetFocusedRowCellValue("DE_NO")?.ToString();
            string authorized = gridView2.GetFocusedRowCellValue("AUTHORIZED")?.ToString();
            string protGuid = gridView2.GetFocusedRowCellValue("PROT_Guid")?.ToString();
            UpdateEntry(id, value.ToString(), authorized, protGuid);
        }

        //private void repositoryItemCheckEdit2_CheckStateChanged(object sender, EventArgs e)
        //{
        //    CheckEdit editor = sender as CheckEdit;
        //    CheckState state = editor.CheckState;

        //    bool isChecked = editor.Checked;

        //    int id = Convert.ToInt32(gridView2.GetFocusedRowCellValue("ID"));
        //    string deNo = gridView2.GetFocusedRowCellValue("DE_NO")?.ToString();
        //    //string authorized = gridView2.GetFocusedRowCellValue("AUTHORIZED")?.ToString();
        //    string protGuid = gridView2.GetFocusedRowCellValue("PROT_Guid")?.ToString();
        //    UpdateEntry(id, deNo, isChecked.ToString(), protGuid);
        //}

        private void hyperlinkLabelControl2_Click(object sender, EventArgs e)
        {
            gridControl3.ShowPrintPreview();
        }
        private void CheckEdit_CheckStateChanged(object sender, EventArgs e)
        {
            CheckEdit checkEdit = sender as CheckEdit;

            if (checkEdit != null)
            {
                int rowHandle = gridView3.FocusedRowHandle;



                string columnHandle = gridView3.FocusedColumn.FieldName;

                string protUser = gridView3.GetRowCellValue(rowHandle, "PROT_User")?.ToString();
                string protGuid = gridView3.GetRowCellValue(rowHandle, "PROT_Guid")?.ToString();
                string depotname = columnHandle;
                string authorized = gridView3.GetRowCellValue(rowHandle, depotname)?.ToString();
                string deno = GetDE_NoByIntitule(depotname).ToString();
                int id = GetIDByDeno(deno, protGuid);
                string val;

                bool isChecked = checkEdit.Checked;
                if (isChecked)
                {
                    val = "1";
                }
                else
                {
                    val = "0";

                }

                // MessageBox.Show($"PROT_User: {protUser}, Checked: {isChecked}, Dépôt: {depotname}");
                UpdateEntry(id, deno, val, protGuid);
            }
        }

        private int GetDE_NoByIntitule(string deIntitule)
        {
            int deNo = -1; // Valeur par défaut au cas où aucun résultat n'est trouvé

            // Connexion à la base de données
            string cs = $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";

            using (SqlConnection connection = new SqlConnection(cs))
            {
                // Requête SQL pour récupérer la valeur DE_No en fonction de DE_Intitule
                string query = "SELECT DE_No FROM F_DEPOT WHERE DE_Intitule = @DE_Intitule";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Ajouter le paramètre pour éviter les attaques par injection SQL
                    command.Parameters.AddWithValue("@DE_Intitule", deIntitule);

                    try
                    {
                        connection.Open();
                        // Exécuter la requête et récupérer la valeur de DE_No
                        var result = command.ExecuteScalar();

                        // Si un résultat est trouvé, le convertir en entier
                        if (result != DBNull.Value && result != null)
                        {
                            deNo = Convert.ToInt32(result);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Erreur lors de la récupération de DE_No: " + ex.Message);
                    }
                }
            }

            return deNo; // Retourner la valeur DE_No ou -1 si pas trouvé
        }
        private int GetIDByDeno(string deno, string protguid)
        {
            int deNo = -1; // Valeur par défaut au cas où aucun résultat n'est trouvé

            // Connexion à la base de données
            string cs = $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";

            using (SqlConnection connection = new SqlConnection(cs))
            {
                // Requête SQL pour récupérer la valeur DE_No en fonction de DE_Intitule
                string query = "SELECT ID FROM F_DEPOT_DEDIE WHERE DE_NO = @DE_NO AND PROT_Guid = @protguid";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Ajouter le paramètre pour éviter les attaques par injection SQL
                    command.Parameters.AddWithValue("@DE_NO", deno);
                    command.Parameters.AddWithValue("@protguid", protguid);

                    try
                    {
                        connection.Open();
                        // Exécuter la requête et récupérer la valeur de DE_No
                        var result = command.ExecuteScalar();

                        // Si un résultat est trouvé, le convertir en entier
                        if (result != DBNull.Value && result != null)
                        {
                            deNo = Convert.ToInt32(result);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Erreur lors de la récupération de DE_No: " + ex.Message);
                    }
                }
            }

            return deNo; // Retourner la valeur DE_No ou -1 si pas trouvé
        }

        private void gridView3_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            //if (e.Column != null)
            //{
            //    string fieldName = e.Column.FieldName;

            //    // Utilisation de switch pour matcher des mots-clés
            //    switch (fieldName)
            //    {
            //        case string name when name.Contains("AMBOHIMANGAKELY"):
            //            e.Appearance.BackColor = Color.LightBlue;
            //            e.Appearance.ForeColor = Color.DarkBlue;
            //            break;

            //        case string name when name.Contains("ANALAKELY"):
            //            e.Appearance.BackColor = Color.LightGreen;
            //            e.Appearance.ForeColor = Color.Black;
            //            break;

            //        case string name when name.Contains("ANTANIMORA"):
            //            e.Appearance.BackColor = Color.LightYellow;
            //            e.Appearance.ForeColor = Color.DarkRed;
            //            break;

            //        case string name when name.Contains("DIEGO"):
            //            e.Appearance.BackColor = Color.Coral;
            //            e.Appearance.ForeColor = Color.DarkRed;
            //            break;

            //        case string name when name.Contains("IMERINTSIATOSIKA"):
            //            e.Appearance.BackColor = Color.LightSteelBlue;
            //            e.Appearance.ForeColor = Color.DarkRed;
            //            break;

            //        case string name when name.Contains("MAHITSY"):
            //            e.Appearance.BackColor = Color.LightYellow;
            //            e.Appearance.ForeColor = Color.DarkRed;
            //            break;

            //        case string name when name.Contains("TMM"):
            //            e.Appearance.BackColor = Color.Plum;
            //            e.Appearance.ForeColor = Color.DarkRed;
            //            break;

            //        default:
            //            // Optionnel : apparence par défaut
            //            e.Appearance.BackColor = Color.White;
            //            e.Appearance.ForeColor = Color.Black;
            //            break;
            //    }
            //}
        }

        private void hyperlinkLabelControl3_Click(object sender, EventArgs e)
        {
            FrmNewUser frmNewuser = new FrmNewUser();
            frmNewuser.ShowDialog();
        }
        private int hoveredRowHandle = GridControl.InvalidRowHandle;
        private void gridView3_MouseMove(object sender, MouseEventArgs e)
        {
            GridView view = sender as GridView;
            GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            if (hitInfo.InRow && hitInfo.RowHandle != hoveredRowHandle)
            {
                hoveredRowHandle = hitInfo.RowHandle;
                view.RefreshRow(hitInfo.RowHandle);
            }
        }

        private void gridView3_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle == hoveredRowHandle)
            {
                e.Appearance.BackColor = Color.LightCyan; // Changez la couleur de survol
            }
        }

        private void gridView4_PopupMenuShowing(object sender, DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs e)
        {
            if (e.HitInfo.InRow) // clic droit sur une ligne
            {
                gridView4.FocusedRowHandle = e.HitInfo.RowHandle; // sélectionne la ligne
                contextMenuStrip1.Show(Cursor.Position); // affiche le menu contextuel
            }
        }
        private void SupprimerLigne_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView4.FocusedRowHandle >= 0)
                {
                    DataRowView rowView = gridView4.GetRow(gridView4.FocusedRowHandle) as DataRowView;
                    if (rowView != null)
                    {
                        string cbSession = rowView["cbSession"].ToString(); // Adapter le nom de la clé primaire

                        string leserveur = FrmMdiParent.DataSourceNameValueParent;
                        string connectionString = $"Data Source={leserveur};Initial Catalog=ARBIOCHEM;User ID=Dev;Password=1234;TrustServerCertificate=True";

                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            SqlCommand cmd = new SqlCommand("DELETE FROM cbUserSession WHERE cbSession = @cbSession", connection);
                            cmd.Parameters.AddWithValue("@cbSession", cbSession);
                            cmd.ExecuteNonQuery();
                        }

                        // Supprime la ligne de la vue
                        gridView4.DeleteRow(gridView4.FocusedRowHandle);
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void hyperlinkLabelControl4_Click(object sender, EventArgs e)
        {
            LoadcbUserSession();
        }

        private void gridView4_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }
        int hotTrackRow = DevExpress.XtraGrid.GridControl.InvalidRowHandle;

        private void gridView3_MouseMove_1(object sender, MouseEventArgs e)
        {
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(e.Location);

            if (info.InRowCell)
            {
                if (hotTrackRow != info.RowHandle)
                {
                    int oldHotTrack = hotTrackRow;
                    hotTrackRow = info.RowHandle;
                    view.RefreshRow(oldHotTrack);
                    view.RefreshRow(hotTrackRow);
                }
            }
            else
            {
                if (hotTrackRow != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
                {
                    int oldHotTrack = hotTrackRow;
                    hotTrackRow = DevExpress.XtraGrid.GridControl.InvalidRowHandle;
                    view.RefreshRow(oldHotTrack);
                }
            }
        }

        private void gridView3_RowStyle_1(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle == hotTrackRow)
            {
                e.Appearance.BackColor = Color.LightYellow;  // couleur de survol
                e.Appearance.ForeColor = Color.Black;
            }
        }
    }
}

