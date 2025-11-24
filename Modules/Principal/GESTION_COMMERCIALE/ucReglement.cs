using DevExpress.XtraEditors;
using DevExpress.XtraPivotGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.GESTION_COMMERCIALE
{
    public partial class ucReglement : DevExpress.XtraEditors.XtraUserControl
    {
        private static ucReglement _instance;
        public static ucReglement Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucReglement();
                return _instance;
            }
        }
        public ucReglement()
        {
            InitializeComponent();
            LoadReglement();
        }

        private void LoadReglement()
        {

        }
        //private Task ChargerPivot()
        //{
        //    try
        //    {
        //        //splashScreenManager1.ShowWaitForm();
        //        //System.Data.DataTable table = GetData();
        //        if (table.Rows.Count == 0)
        //        {
        //            MessageBox.Show("Aucune donnée disponible pour afficher.");
        //            return Task.CompletedTask;
        //        }

        //        // Effacer les champs précédents pour éviter les doublons
        //        pivotGridControl1.Fields.Clear();

        //        // Configuration des lignes (Row Fields)
        //        PivotGridField fieldCat = new PivotGridField("Catalog", PivotArea.RowArea);
        //        fieldCat.Caption = "Site";
        //        pivotGridControl1.Fields.Add(fieldCat);

        //        PivotGridField fieldArRef = new PivotGridField("ar_ref", PivotArea.RowArea);
        //        fieldArRef.Caption = "Référence";
        //        pivotGridControl1.Fields.Add(fieldArRef);

        //        PivotGridField fieldDesignation = new PivotGridField("DESIGNATION", PivotArea.RowArea);
        //        fieldDesignation.Caption = "Désignation";
        //        pivotGridControl1.Fields.Add(fieldDesignation);

        //        PivotGridField fieldCategorie = new PivotGridField("CATEGORIE", PivotArea.RowArea);
        //        fieldDesignation.Caption = "Catégorie tarifaire";
        //        pivotGridControl1.Fields.Add(fieldCategorie);

        //        //// Configuration des colonnes (Column Fields)
        //        PivotGridField fieldHDate = new PivotGridField("HDate", PivotArea.ColumnArea)
        //        {
        //            GroupInterval = PivotGroupInterval.Date
        //        };
        //        fieldHDate.Caption = "Date";
        //        pivotGridControl1.Fields.Add(fieldHDate);


        //        // Configuration des valeurs (Data Fields)
        //        PivotGridField fieldAncienPrix = new PivotGridField("ANCIEN_PRIX", PivotArea.DataArea);
        //        fieldAncienPrix.Caption = "Ancien Prix";
        //        fieldAncienPrix.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
        //        fieldAncienPrix.CellFormat.FormatString = "N2"; // Format numérique à 2 décimales
        //        pivotGridControl1.Fields.Add(fieldAncienPrix);

        //        PivotGridField fieldPrixActuel = new PivotGridField("PRIX_ACTUEL", PivotArea.DataArea);
        //        fieldPrixActuel.Caption = "Prix Actuel";
        //        fieldPrixActuel.CellFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
        //        fieldPrixActuel.CellFormat.FormatString = "N2"; // Format numérique à 2 décimales
        //        pivotGridControl1.Fields.Add(fieldPrixActuel);

        //        // Affecter le DataTable comme source de données
        //        pivotGridControl1.DataSource = table;

        //        // Supprimer les grands totaux horizontaux
        //        pivotGridControl1.OptionsView.ShowColumnGrandTotals = false;

        //        // Configuration d'options supplémentaires si nécessaire
        //        pivotGridControl1.OptionsData.DataProcessingEngine = PivotDataProcessingEngine.Optimized;
        //        pivotGridControl1.BestFitRowArea();
        //        //splashScreenManager1.CloseWaitForm();
        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    finally
        //    {
        //        //SplashScreenManager.CloseForm(false);
        //    }

        //    return Task.CompletedTask;
        //}
        //private System.Data.DataTable GetData()
        //{
        //    //var serveursEtBases = GetServeursEtBases();
        //    System.Data.DataTable table = new System.Data.DataTable();

        //    foreach (var (serveur, baseDeDonnees) in serveursEtBases)
        //    {
        //        try
        //        {
        //            //if (baseDeDonnees != "AMBOHIMANGAKELY_OLD") ; break;
        //            string connectionString = $"Server={serveur};Database={baseDeDonnees};User ID=Dev;Password=1234;TrustServerCertificate=True";
        //            using (SqlConnection conn = new SqlConnection(connectionString))
        //            {
        //                conn.OpenAsync();

        //                string query = @"SELECT DISTINCT
        //                          dbo.F_ARTCLIENT_DEMANDE.DataSource,
        //                          dbo.F_ARTCLIENT_DEMANDE.Catalog,
        //                          dbo.F_ARTCLIENT_DEMANDE.ar_ref,
        //                          dbo.F_ARTCLIENT_DEMANDE.DESIGNATION,
        //                          dbo.F_ARTCLIENT_DEMANDE.CATEGORIE,   
        //                          dbo.F_ARTCLIENT_DEMANDE.ANCIEN_PRIX,
        //                          dbo.F_ARTCLIENT.AC_PrixVen AS PRIX_ACTUEL,
  
        //                          dbo.F_ARTCLIENT_DEMANDE.UserName,
        //                          dbo.F_ARTCLIENT_DEMANDE.UserRole,
        //                          dbo.F_ARTCLIENT_DEMANDE.HDate,
        //                          dbo.F_ARTCLIENT_DEMANDE.ENVOIMAIL,
        //                          dbo.F_ARTCLIENT_DEMANDE.UPDATED
        //                        FROM
        //                          dbo.F_ARTCLIENT
        //                          INNER JOIN dbo.F_ARTCLIENT_DEMANDE ON (dbo.F_ARTCLIENT.AR_Ref = dbo.F_ARTCLIENT_DEMANDE.ar_ref
        //                          and (dbo.F_ARTCLIENT.AC_PrixVen = dbo.F_ARTCLIENT_DEMANDE.NOUVEAU_PRIX))";
        //                using (SqlCommand cmd = new SqlCommand(query, conn))
        //                {
        //                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
        //                    adapter.Fill(table);
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            //MessageBox.Show($"Erreur avec le serveur {serveur} et la base {baseDeDonnees}: {ex.Message}");
        //        }
        //    }
        //    return table;
        //}
    }
}
