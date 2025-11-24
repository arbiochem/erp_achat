using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraPivotGrid;
using DevExpress.XtraSplashScreen;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DSI
{
    public partial class ucUpdatePrix : DevExpress.XtraEditors.XtraUserControl
    {
        private static ucUpdatePrix _instance;
        public static ucUpdatePrix Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucUpdatePrix();
                return _instance;
            }
        }

        public ucUpdatePrix()
        {
            InitializeComponent();
        }

        private void officeNavigationBar1_Click(object sender, EventArgs e)
        {

        }
        private async Task IterateBdAsync()
        {
            var serveursEtBases = GetServeursEtBases();
            System.Data.DataTable dataTable = new System.Data.DataTable();

            gridControl1.DataSource = null;
            gridView1.Columns.Clear();

            foreach (var (serveur, baseDeDonnees) in serveursEtBases)
            {
                try
                {
                    //if (baseDeDonnees != "AMBOHIMANGAKELY_OLD") ; break;
                    string connectionString = $"Server={serveur};Database={baseDeDonnees};User ID=Dev;Password=1234;TrustServerCertificate=True";
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        await conn.OpenAsync();

                        string query = @"SELECT 
                                          ID,
                                          DataSource,
                                          Catalog,
                                          ar_ref,
                                          DESIGNATION,
                                          CATEGORIE,
                                          ANCIEN_PRIX,
                                          NOUVEAU_PRIX,
                                          UserName,
                                          UserRole,
                                          HDate,
                                          ENVOIMAIL,
                                          UPDATED
                                        FROM 
                                          dbo.F_ARTCLIENT_DEMANDE";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"Erreur avec le serveur {serveur} et la base {baseDeDonnees}: {ex.Message}");
                }
            }
            gridControl1.DataSource = dataTable;    
        }
        private List<(string serveur, string baseDeDonnees)> GetServeursEtBases()
        {
            string DataSource = FrmMdiParent.DataSourceNameValueParent;
            var serveursEtBases = new List<(string serveur, string baseDeDonnees)>();

            string query = "SELECT Site, connex FROM T_Server"; // La table T_Server contient les informations nécessaires

            // Remplacer par votre chaîne de connexion de base pour récupérer les données
            //string connectionString = "votre_chaine_de_connexion"; 
            string connectionString = $"Data Source={DataSource};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string serveur = reader["connex"].ToString();
                    string baseDeDonnees = reader["Site"].ToString();

                    serveursEtBases.Add((serveur, baseDeDonnees)); // Ajouter le serveur et la base de données à la liste
                }
            }

            return serveursEtBases;
        }


        private void UpdateF_ARTCLIENTFromGrid(int row)
        {
            if (row < 0 || row >= gridView1.RowCount)
            {
                MessageBox.Show("Index de ligne invalide.");
                return;
            }

            // Récupération des données de la ligne spécifiée
            var dataRow = gridView1.GetDataRow(row);
            if (dataRow == null)
            {
                MessageBox.Show("Impossible de récupérer les données de la ligne.");
                return;
            }

            string datasource = dataRow["DataSource"].ToString();
            string bdd = dataRow["Catalog"].ToString();
            string arRef = dataRow["ar_ref"].ToString();
            string categorie = dataRow["CATEGORIE"].ToString();
            decimal nouveauPrix = Convert.ToDecimal(dataRow["NOUVEAU_PRIX"]);
            int updated = Convert.ToInt16(dataRow["UPDATED"]);

            string connectionString = $"Server={datasource};Database={bdd};User ID=Dev;Password=1234;TrustServerCertificate=True";

            // Requêtes SQL
            string updateF_ARTCLIENTQuery = @"UPDATE F_ARTCLIENT
                        SET AC_PrixVen = @NouveauPrix
                        FROM F_ARTCLIENT
                        LEFT JOIN p_cattarif c ON F_ARTCLIENT.ac_categorie = c.cbIndice
                        WHERE F_ARTCLIENT.ar_ref = @AR_Ref AND c.CT_Intitule = @CatTarif;";

            string updateF_ARTCLIENT_DEMANDEQuery = @"UPDATE f
                        SET f.UPDATED = 1
                        FROM F_ARTCLIENT_DEMANDE f                               
                        WHERE f.ar_ref = @AR_Ref AND f.CATEGORIE = @CatTarif;";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // Mise à jour de F_ARTCLIENT
                    using (SqlCommand cmdUpdateF_ARTCLIENT = new SqlCommand(updateF_ARTCLIENTQuery, conn))
                    {
                        cmdUpdateF_ARTCLIENT.Parameters.AddWithValue("@NouveauPrix", nouveauPrix);
                        cmdUpdateF_ARTCLIENT.Parameters.AddWithValue("@AR_Ref", arRef);
                        cmdUpdateF_ARTCLIENT.Parameters.AddWithValue("@CatTarif", categorie);

                        if (updated == 0)
                        {
                            int rowsAffected = cmdUpdateF_ARTCLIENT.ExecuteNonQuery();
                            if (rowsAffected > 0)
                            {
                                // Mise à jour de F_ARTCLIENT_DEMANDE
                                using (SqlCommand cmdUpdateF_ARTCLIENT_DEMANDE = new SqlCommand(updateF_ARTCLIENT_DEMANDEQuery, conn))
                                {
                                    cmdUpdateF_ARTCLIENT_DEMANDE.Parameters.AddWithValue("@AR_Ref", arRef);
                                    cmdUpdateF_ARTCLIENT_DEMANDE.Parameters.AddWithValue("@CatTarif", categorie);
                                    int demandeRowsAffected = cmdUpdateF_ARTCLIENT_DEMANDE.ExecuteNonQuery();

                                    // Afficher le résultat
                                    MessageBox.Show($"Mise à jour réussie!");
                                }

                                // Mettre à jour la colonne UPDATED dans GridControl
                                dataRow["UPDATED"] = 1;
                            }
                            else
                            {
                                MessageBox.Show("Aucune mise à jour dans F_ARTCLIENT.");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Cette ligne a déjà été mise à jour.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Gérer les erreurs
                MessageBox.Show($"Erreur lors de la mise à jour : {ex.Message}");
            }

            // Rafraîchir l'affichage de la grille après la mise à jour
            gridView1.RefreshData();
        }
        private System.Data.DataTable GetData()
        {
            var serveursEtBases = GetServeursEtBases();
            System.Data.DataTable table = new System.Data.DataTable();

            foreach (var (serveur, baseDeDonnees) in serveursEtBases)
            {
                try
                {
                    //if (baseDeDonnees != "AMBOHIMANGAKELY_OLD") ; break;
                    string connectionString = $"Server={serveur};Database={baseDeDonnees};User ID=Dev;Password=1234;TrustServerCertificate=True";
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.OpenAsync();

                        string query = @"SELECT DISTINCT
                                  dbo.F_ARTCLIENT_DEMANDE.DataSource,
                                  dbo.F_ARTCLIENT_DEMANDE.Catalog,
                                  dbo.F_ARTCLIENT_DEMANDE.ar_ref,
                                  dbo.F_ARTCLIENT_DEMANDE.DESIGNATION,
                                  dbo.F_ARTCLIENT_DEMANDE.CATEGORIE,   
                                  dbo.F_ARTCLIENT_DEMANDE.ANCIEN_PRIX,
                                  dbo.F_ARTCLIENT.AC_PrixVen AS PRIX_ACTUEL,
  
                                  dbo.F_ARTCLIENT_DEMANDE.UserName,
                                  dbo.F_ARTCLIENT_DEMANDE.UserRole,
                                  dbo.F_ARTCLIENT_DEMANDE.HDate,
                                  dbo.F_ARTCLIENT_DEMANDE.ENVOIMAIL,
                                  dbo.F_ARTCLIENT_DEMANDE.UPDATED
                                FROM
                                  dbo.F_ARTCLIENT
                                  INNER JOIN dbo.F_ARTCLIENT_DEMANDE ON (dbo.F_ARTCLIENT.AR_Ref = dbo.F_ARTCLIENT_DEMANDE.ar_ref
                                  and (dbo.F_ARTCLIENT.AC_PrixVen = dbo.F_ARTCLIENT_DEMANDE.NOUVEAU_PRIX))";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(table);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"Erreur avec le serveur {serveur} et la base {baseDeDonnees}: {ex.Message}");
                }
            }
            return table;
        }
        private async void hyperlinkLabelControl4_Click(object sender, EventArgs e)
        {
            await IterateBdAsync();
            await ChargerPivot();
        }
        private Task ChargerPivot()
        {
            try
            {
                //splashScreenManager1.ShowWaitForm();
                System.Data.DataTable table = GetData();
                if (table.Rows.Count == 0)
                {
                    MessageBox.Show("Aucune donnée disponible pour afficher.");
                    return Task.CompletedTask;
                }

                // Effacer les champs précédents pour éviter les doublons
                pivotGridControl1.Fields.Clear();

                // Configuration des lignes (Row Fields)
                PivotGridField fieldCat = new PivotGridField("Catalog", PivotArea.RowArea);
                fieldCat.Caption = "Site";
                pivotGridControl1.Fields.Add(fieldCat);

                PivotGridField fieldArRef = new PivotGridField("ar_ref", PivotArea.RowArea);
                fieldArRef.Caption = "Référence";
                pivotGridControl1.Fields.Add(fieldArRef);

                PivotGridField fieldDesignation = new PivotGridField("DESIGNATION", PivotArea.RowArea);
                fieldDesignation.Caption = "Désignation";
                pivotGridControl1.Fields.Add(fieldDesignation);

                PivotGridField fieldCategorie = new PivotGridField("CATEGORIE", PivotArea.RowArea);
                fieldDesignation.Caption = "Catégorie tarifaire";
                pivotGridControl1.Fields.Add(fieldCategorie);

                //// Configuration des colonnes (Column Fields)
                PivotGridField fieldHDate = new PivotGridField("HDate", PivotArea.ColumnArea)
                {
                    GroupInterval = PivotGroupInterval.Date
                };
                fieldHDate.Caption = "Date";
                pivotGridControl1.Fields.Add(fieldHDate);

                
                // Configuration des valeurs (Data Fields)
                PivotGridField fieldAncienPrix = new PivotGridField("ANCIEN_PRIX", PivotArea.DataArea);
                fieldAncienPrix.Caption = "Ancien Prix";
                fieldAncienPrix.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                fieldAncienPrix.CellFormat.FormatString = "N2"; // Format numérique à 2 décimales
                pivotGridControl1.Fields.Add(fieldAncienPrix);

                PivotGridField fieldPrixActuel = new PivotGridField("PRIX_ACTUEL", PivotArea.DataArea);
                fieldPrixActuel.Caption = "Prix Actuel";
                fieldPrixActuel.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                fieldPrixActuel.CellFormat.FormatString = "N2"; // Format numérique à 2 décimales
                pivotGridControl1.Fields.Add(fieldPrixActuel);

                // Affecter le DataTable comme source de données
                pivotGridControl1.DataSource = table;

                // Supprimer les grands totaux horizontaux
                pivotGridControl1.OptionsView.ShowColumnGrandTotals = false;

                // Configuration d'options supplémentaires si nécessaire
                pivotGridControl1.OptionsData.DataProcessingEngine = PivotDataProcessingEngine.Optimized;
                pivotGridControl1.BestFitRowArea();
                //splashScreenManager1.CloseWaitForm();
            }

            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //SplashScreenManager.CloseForm(false);
            }

            return Task.CompletedTask;
        }

        private void hyperlinkLabelControl5_Click(object sender, EventArgs e)
        {
            pivotGridControl1.ShowPrintPreview();
        }
               

        private async void demandeEnAttenteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            await IterateBdAsync();
            gridView1.Columns["ID"].VisibleIndex = -1;
            gridView1.BestFitColumns();
            //for (int i = 0; i < gridView1.Columns.Count; i++)
            //{
            //    gridView1.Columns[i].OptionsColumn.ReadOnly = true;
            //}
            gridView1.ActiveFilter.Clear(); // Effacer tous les filtres existants
            gridView1.ActiveFilter.Add(gridView1.Columns["UPDATED"], new ColumnFilterInfo("[UPDATED] = False"));
            gridView1.Columns["ANCIEN_PRIX"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gridView1.Columns["ANCIEN_PRIX"].DisplayFormat.FormatString = "N2";
            gridView1.Columns["NOUVEAU_PRIX"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gridView1.Columns["NOUVEAU_PRIX"].DisplayFormat.FormatString = "N2";

            splashScreenManager1.CloseWaitForm();
        }

        private void exporterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            gridControl1.ShowPrintPreview();
        }

        private void appliquerLesNouveauxPrixToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(
                splashFormType: typeof(WaitForm2),
                useFadeIn: false,
                useFadeOut: false);
            int v = 0;
            int j = gridView1.RowCount;
            int l = 0;
            int g = 0;
            SplashScreenManager.Default.SendCommand(WaitForm2.WaitFormCommand.SetProgressBar, j);
            for (int i = 0; i < gridView1.RowCount; i++)
            {
                v++;
                l = (i * 100) / j;
                g = (i * 1);

                SplashScreenManager.Default.SendCommand(WaitForm2.WaitFormCommand.SetProgress, v);
                SplashScreenManager.Default.SetWaitFormDescription(g.ToString() + "/" + j.ToString());


                int rowHandle = gridView1.GetRowHandle(i);
                bool isUpdated = Convert.ToBoolean(gridView1.GetRowCellValue(i, "UPDATED"));
                if (gridView1.IsDataRow(rowHandle)) // Vérifier que c'est une ligne de données
                {
                    if (!isUpdated)
                    {
                        UpdateF_ARTCLIENTFromGrid(i);
                    }   
                }
            }
            SplashScreenManager.CloseForm(false);
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // Récupérer la ligne actuellement sélectionnée
            int focusedRowHandle = gridView1.FocusedRowHandle;

            if (focusedRowHandle >= 0)
            {
                // Récupérer les valeurs des colonnes
                string dataSourceValue = gridView1.GetRowCellValue(focusedRowHandle, "DataSource")?.ToString();
                string catalogValue = gridView1.GetRowCellValue(focusedRowHandle, "Catalog")?.ToString();
                string idValue = gridView1.GetRowCellValue(focusedRowHandle, "ID")?.ToString();

                // Vérifier que toutes les valeurs sont récupérées
                if (!string.IsNullOrEmpty(dataSourceValue) &&
                    !string.IsNullOrEmpty(catalogValue) &&
                    !string.IsNullOrEmpty(idValue))
                {
                    // Appeler la méthode pour supprimer l'enregistrement de la base de données
                    DeleteRecordFromDatabase(dataSourceValue, catalogValue, idValue);

                    
                }
                else
                {
                    MessageBox.Show("Impossible de récupérer toutes les valeurs nécessaires pour supprimer l'enregistrement.");
                }
            }
            else
            {
                MessageBox.Show("Aucune ligne sélectionnée !");
            }
        }

        
        private void gridView1_PopupMenuShowing(object sender, DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs e)
        {
            if (e.HitInfo.InRow)
            {
                System.Drawing.Point p2 = Control.MousePosition;
                this.popupMenu1.ShowPopup(p2);
            }
        }
        private void DeleteRecordFromDatabase(string dataSource, string catalog, string idRow)
        {
            GridView gridView = gridControl1.MainView as GridView; // Votre GridView
            if (gridView == null || gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("Aucune ligne sélectionnée !");
                return;
            }

            string connectionString = $"Data Source={dataSource};Initial Catalog={catalog};User ID=Dev;Password=1234;TrustServerCertificate=True";
            // Requête SQL pour supprimer l'enregistrement
            string deleteQuery = "DELETE FROM F_ARTCLIENT_DEMANDE WHERE ID = @ID";

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", idRow);
                        
                        int rowsAffected = cmd.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Enregistrement supprimé avec succès !");
                            gridView.DeleteRow(gridView.FocusedRowHandle); // Supprimer la ligne du GridView
                            gridView.RefreshData();
                        }
                        else
                        {
                            MessageBox.Show("Aucun enregistrement correspondant trouvé.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}");
            }
        }

        private void gridView1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)  // Lignes paires
            {
                e.Appearance.BackColor = Color.WhiteSmoke;
            }
            else  // Lignes impaires
            {
                e.Appearance.BackColor = Color.White;
            }
        }

        private async void historiquesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            await AfficherHistorique();
            splashScreenManager1.CloseWaitForm();
        }
        private async Task AfficherHistorique()
        {
            var serveursEtBases = GetServeursEtBases();
            System.Data.DataTable dataTable = new System.Data.DataTable();

            gridControl1.DataSource = null;
            gridView1.Columns.Clear();

            foreach (var (serveur, baseDeDonnees) in serveursEtBases)
            {
                try
                {
                    //if (baseDeDonnees != "AMBOHIMANGAKELY_OLD") ; break;
                    string connectionString = $"Server={serveur};Database={baseDeDonnees};User ID=Dev;Password=1234;TrustServerCertificate=True";
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        await conn.OpenAsync();

                        string query = @"SELECT *
                                        FROM 
                                          dbo.F_ARTCLIENT_History";
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"Erreur avec le serveur {serveur} et la base {baseDeDonnees}: {ex.Message}");
                }
            }
            gridControl1.DataSource = dataTable;
        }
    }
}






    
