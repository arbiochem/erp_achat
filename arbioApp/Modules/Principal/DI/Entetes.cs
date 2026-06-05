using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Models;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;


namespace arbioApp.Modules.Principal.DI
{
    public class Entetes
    {
        public static int rownum = 0;
        private static DataTable dataTable;

        private static string DbPrincipale => ucDocuments.dbNamePrincipale;
        private static string ServerIpPrincipale => ucDocuments.serverIpPrincipale;

        // Base contenant F_DOCLIGNE (peut être différente de DbPrincipale)
        // Modifiez "TRANSIT" par le nom retourné par :
        // SELECT TABLE_CATALOG FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'F_DOCLIGNE'

       
        private static string connectionString =>
            $"Server={ServerIpPrincipale};Database={DbPrincipale};" +
            $"User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=240;";

        private static string connectionStrings =>
            $"Server={ServerIpPrincipale};Database=TRANSIT;" +
            $"User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=120;";

        // ════════════════════════════════════════════════════════════════
        //  AfficherEntetes  (grille simple, inchangée)
        // ════════════════════════════════════════════════════════════════
        public static void AfficherEntetes(GridControl gc, int achattype, BindingSource bs)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "SELECT * FROM dbo.ACHAT_ENTETE ORDER BY DO_Date DESC";
                    try
                    {
                        using (SqlCommand cmd = new SqlCommand(query, connection))
                        {
                            if (connection.State != ConnectionState.Open)
                                connection.Open();

                            using (SqlDataAdapter localAdapter = new SqlDataAdapter(cmd))
                            {
                                dataTable = new DataTable();
                                localAdapter.Fill(dataTable);
                                rownum = dataTable.Rows.Count;

                                if (rownum == 0)
                                {
                                    bs.DataSource = null;
                                    gc.DataSource = null;
                                }
                                else
                                {
                                    bs.DataSource = dataTable;
                                    gc.DataSource = bs;
                                }
                            }
                        }
                    }
                    catch (SqlException ex) { MessageBox.Show("Erreur SQL : " + ex.Message); }
                    catch (Exception ex) { MessageBox.Show("Erreur : " + ex.Message); }
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}",
                                "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);*/
            }
        }



        // ════════════════════════════════════════════════════════════════
        //  AfficherEntetes_achat  (avec Master-Detail sur gc2 et gc3)
        // ════════════════════════════════════════════════════════════════
        public static void AfficherEntetes_achat(
            GridControl gc, GridControl gc1, GridControl gc2, GridControl gc3,
            int achattype, BindingSource bs)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = $@"
                    WITH LignesGroupees AS (
                        SELECT
                            Do_Piece,
                            AR_Ref,
                            DL_Design,
                            SUM(DL_Qte)                                                AS Total_Qte,
                            SUM(ISNULL(TRY_CAST(QteLivre AS DECIMAL(18,2)), 0))       AS QteLivre,
                            MAX(DL_Qte)                                                AS DL_Qte,
                            SUM(DL_PrixRU)                                             AS Total_PrixRU,
                            SUM(DL_MontantTTC)                                         AS Total_MontantTTC
                        FROM dbo.F_DOCLIGNE
                        GROUP BY Do_Piece, AR_Ref, DL_Design
                    ),
                    LignesAvecProduit AS (
                        SELECT 
                            ent.DO_Piece                                               AS DO_Piece_Entete,
                            MIN(lg.AR_Ref)                                             AS AR_Ref,
                            MIN(lg.DL_Design)                                          AS DL_Design,
                            SUM(lg.QteLivre)                                           AS QteLivre,
                            MAX(lg.DL_Qte)                                             AS DL_Qte,
                            SUM(lg.Total_PrixRU)                                       AS Total_PrixRU,
                            SUM(lg.Total_MontantTTC)                                   AS Total_MontantTTC
                        FROM dbo.ACHAT_ENTETE ent
                        INNER JOIN LignesGroupees lg 
                            ON REPLACE(lg.Do_Piece, '_', '') = ent.DO_Piece
                            OR lg.Do_Piece = ent.PRODUIT
                        GROUP BY ent.DO_Piece  -- ✅ groupé par DO_Piece seulement
                    )
                    SELECT
                        ent.DO_Piece,
                        ent.CT_Intitule,
                        lp.AR_Ref,
                        lp.DL_Design,
                        CASE 
                            WHEN lp.DL_Qte != lp.QteLivre THEN FORMAT(lp.DL_Qte + lp.QteLivre, 'N2')
                            ELSE FORMAT(lp.DL_Qte, 'N2')
                        END AS DL_Qte,
                        FORMAT(lp.QteLivre, 'N2')                                      AS DL_QteLivre,
                        CASE 
                            WHEN lp.DL_Qte = lp.QteLivre THEN FORMAT(lp.DL_Qte - lp.QteLivre, 'N2')
                            ELSE FORMAT(lp.DL_Qte, 'N2')
                        END AS DL_QteRestant,
                        CAST(FORMAT(lp.Total_PrixRU, 'N2') AS VARCHAR(50))
                            + ' ' + ent.D_Intitule                                     AS DL_PrixRU,
                        CAST(FORMAT(lp.Total_MontantTTC / NULLIF(ent.DO_Cours, 0), 'N2') AS VARCHAR(50))
                            + ' ' + ent.D_Intitule                                     AS TotalTTC,
                        ISNULL(CAST(ent.DO_Cloture AS INT), 0)                         AS DO_Cloture,
                        ent.CO_No,
                        ent.DE_No,
                        ent.DO_TIERS,
                        ent.DO_Statut,
                        ent.DO_Expedit,
                        ent.DO_Coord01,
                        ent.DO_Ref,
                        ent.DO_DateLivr,
                        ent.DO_Date,
                        ent.DO_DateExpedition,
                        ent.A_TYPE,
                        ent.DO_CodeTaxe1,
                        ent.DO_Taxe1,
                        ent.DO_Type,
                        ent.DO_Imprim,
                        ent.DO_Reliquat
                    FROM dbo.ACHAT_ENTETE AS ent
                    INNER JOIN LignesAvecProduit AS lp ON REPLACE(lp.DO_Piece_Entete, '_', '') = ent.DO_Piece
                    ORDER BY ent.DO_Piece ASC";

                    try
                    {
                        using (SqlCommand cmd = new SqlCommand(query, connection))
                        {
                            if (connection.State != ConnectionState.Open)
                                connection.Open();

                            using (SqlDataAdapter localAdapter = new SqlDataAdapter(cmd))
                            {
                                // ── 1. Charger les données brutes ─────────────────────
                                dataTable = new DataTable();
                                localAdapter.Fill(dataTable);
                                rownum = dataTable.Rows.Count;

                                // ── 2. Créer les DataSets ─────────────────────────────
                                //       gc  → APA / BC   (grille simple)
                                //       gc1 → AFA        (grille simple)
                                //       gc2 → ABR non clôturés  (Master-Detail)
                                //       gc3 → ABR clôturés      (Master-Detail)

                                // Tables plates pour gc / gc1
                                DataTable dtAPA = dataTable.Clone();
                                DataTable dtAutre = dataTable.Clone();
                                DataTable dtCloture = dataTable.Clone();

                                // DataSets Master-Detail pour gc2 / gc3
                                DataSet dsLivre = CreateMasterDetailDataSet(
                                                        "EnteteLivre", "LignesLivre", withAction: true);

                                // ── 3. Répartir les lignes ────────────────────────────
                                foreach (DataRow row in dataTable.Rows)
                                {
                                    string doPiece = row["DO_Piece"]?.ToString() ?? "";
                                    int doCloture = row["DO_Cloture"] == DBNull.Value
                                                        ? 0
                                                        : Convert.ToInt32(row["DO_Cloture"]);

                                    if (doPiece.Contains("APA") || doPiece.Contains("ABC"))
                                    {
                                        dtAPA.ImportRow(row);
                                    }
                                    else if (doPiece.Contains("AFA"))
                                    {
                                        dtAutre.ImportRow(row);
                                    }
                                    else if (doPiece.Contains("ABR"))
                                    {
                                        if (doCloture == 0)
                                            AddRowToDataSet(dsLivre, "EnteteLivre", "LignesLivre", row, action: "Clôturer");
                                        else
                                            dtCloture.ImportRow(row);
                                    }
                                }

                                // ── 4. Colonnes cachées pour grilles simples ──────────
                                string[] colsCachees = {
                                    "DO_Cloture","CO_No","DE_No","DO_TIERS","DO_Statut",
                                    "DO_Expedit","DO_Coord01","DO_Ref","DO_DateLivr",
                                    "DO_Date","DO_DateExpedition","A_TYPE","DO_CodeTaxe1",
                                    "DO_Taxe1","DO_Type","DO_Imprim","DO_Reliquat",
                                    "AR_Ref","DL_Design","DL_Qte","DL_QteLivre","DL_QteRestant"
                                };


                                foreach (DataTable dt in new[] { dtAPA, dtAutre,dtCloture })
                                {
                                    foreach (string col in colsCachees)
                                        if (dt.Columns.Contains(col))
                                            dt.Columns[col].ColumnMapping = MappingType.Hidden;

                                    if (dt.Columns.Contains("DO_Piece"))
                                        dt.Columns["DO_Piece"].Caption = "Numéro pièce";
                                    if (dt.Columns.Contains("CT_Intitule"))
                                        dt.Columns["CT_Intitule"].Caption = "Fournisseur";
                                    if (dt.Columns.Contains("TotalTTC"))
                                        dt.Columns["TotalTTC"].Caption = "Montant TTC";
                                    if (dt.Columns.Contains("DL_PrixRU"))
                                        dt.Columns["DL_PrixRU"].Caption = "Prix unitaire";
                                }

                                // ── 5. Binding des GridControls ───────────────────────

                                // gc → APA/BC  (grille simple)
                                BindingSource bsAPA = new BindingSource { DataSource = dtAPA };
                                gc.DataSource = bsAPA;

                                // gc1 → AFA  (grille simple)
                                BindingSource bsAFA = new BindingSource { DataSource = dtAutre };
                                gc1.DataSource = bsAFA;

                                BindingSource bsCloture = new BindingSource { DataSource = dtCloture };
                                gc3.DataSource = bsCloture;

                                // gc2 → ABR non clôturés  (Master-Detail)
                                BindMasterDetailGrid(
                                    gc2, dsLivre,
                                    "EnteteLivre", "LignesLivre",
                                    "EnteteLivreToLignesLivre",
                                    isLivre: true);

                                // ── 6. BindingSource global ───────────────────────────
                                bs.RaiseListChangedEvents = false;
                                bs.DataSource = null;
                                bs.DataMember = "";
                                if (rownum > 0)
                                    bs.DataSource = dataTable;
                                bs.RaiseListChangedEvents = true;
                            }
                        }
                    }
                    catch (SqlException ex) { MessageBox.Show("Erreur SQL : " + ex.Message); }
                    catch (Exception ex) { MessageBox.Show("Erreur : " + ex.Message); }
                }

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = $@"
                        WITH LignesGroupees AS (
                            SELECT
                                DO_Piece,
                                SUM(ISNULL(TRY_CAST(QteLivre AS DECIMAL(18,2)), 0)) AS QteLivre,
                                MIN(DL_Qte)                                  AS Total_Qte,
                                MIN(AR_Ref)                                  AS AR_Ref,
                                MIN(DL_Design)                               AS DL_Design,
                                SUM(DL_PrixRU)                               AS Total_PrixRU,
                                SUM(DL_MontantTTC)                           AS Total_MontantTTC
                            FROM dbo.F_DOCLIGNE
                            GROUP BY Do_Piece
                        ),
                        LignesAvecProduit AS (
                            SELECT 
                                ent.DO_Piece                                 AS DO_Piece_Entete,
                                MIN(lg.QteLivre)                             AS QteLivre,
                                MIN(lg.Total_Qte)                            AS Total_Qte,
                                MIN(lg.AR_Ref)                               AS AR_Ref,
                                MIN(lg.DL_Design)                            AS DL_Design,
                                SUM(lg.Total_PrixRU)                         AS Total_PrixRU,
                                SUM(lg.Total_MontantTTC)                     AS Total_MontantTTC
                            FROM dbo.ACHAT_ENTETE ent
                            INNER JOIN LignesGroupees lg 
                                ON lg.Do_Piece = ent.DO_Piece
                                OR lg.Do_Piece = ent.PRODUIT
                            GROUP BY ent.DO_Piece
                        )
                        SELECT
                            CASE 
                                WHEN lp.QteLivre > 0 THEN REPLACE(ent.DO_Piece, 'ABR', 'AFA')
                                ELSE ent.DO_Piece
                            END                                              AS DO_Piece,
                            ent.CT_Intitule,
                            lp.AR_Ref,
                            lp.DL_Design,
                            lp.QteLivre,
                            lp.Total_Qte,
                            CAST(FORMAT(lp.Total_PrixRU, 'N2') AS VARCHAR(50))
                                + ' ' + ent.D_Intitule                       AS DL_PrixRU,
                            CAST(FORMAT(lp.Total_MontantTTC / NULLIF(ent.DO_Cours, 0), 'N2') AS VARCHAR(50))
                                + ' ' + ent.D_Intitule                       AS TotalTTC,
                            ISNULL(CAST(ent.DO_Cloture AS INT), 0)           AS DO_Cloture,
                            ent.CO_No,
                            ent.DE_No,
                            ent.DO_TIERS,
                            ent.DO_Statut,
                            ent.DO_Expedit,
                            ent.DO_Coord01,
                            ent.DO_Ref,
                            ent.DO_DateLivr,
                            ent.DO_Date,
                            ent.DO_DateExpedition,
                            ent.A_TYPE,
                            ent.DO_CodeTaxe1,
                            ent.DO_Taxe1,
                            ent.DO_Type,
                            ent.DO_Imprim,
                            ent.DO_Reliquat
                        FROM dbo.ACHAT_ENTETE AS ent
                        INNER JOIN LignesAvecProduit AS lp ON lp.DO_Piece_Entete = ent.DO_Piece
                        WHERE lp.QteLivre != lp.Total_Qte
                        ORDER BY ent.DO_Piece ASC
                    ";

                    try
                    {
                        using (SqlCommand cmd = new SqlCommand(query, connection))
                        {
                            if (connection.State != ConnectionState.Open)
                                connection.Open();

                            using (SqlDataAdapter localAdapter = new SqlDataAdapter(cmd))
                            {
                                // ── 1. Charger les données brutes ─────────────────────
                                dataTable = new DataTable();
                                localAdapter.Fill(dataTable);
                                rownum = dataTable.Rows.Count;

                                DataTable dtAPA = dataTable.Clone();
                                DataTable dtAutre = dataTable.Clone();

                                foreach (DataRow row in dataTable.Rows)
                                {
                                    string doPiece = row["DO_Piece"]?.ToString() ?? "";
                                    //string doPiece = row["DO_Piece"]?.ToString() ?? "";
                                    int doCloture = row["DO_Cloture"] == DBNull.Value
                                                        ? 0
                                                        : Convert.ToInt32(row["DO_Cloture"]);
                                    if (doPiece.Contains("APA") || doPiece.Contains("ABC"))
                                    {
                                        dtAPA.ImportRow(row);
                                    }
                                    else if (doPiece.Contains("AFA"))
                                    {
                                        dtAutre.ImportRow(row);
                                    }
                                }

                                string[] colsCachees = {
                                    "DO_Cloture","CO_No","DE_No","DO_TIERS","DO_Statut",
                                    "DO_Expedit","DO_Coord01","DO_Ref","DO_DateLivr",
                                    "DO_Date","DO_DateExpedition","A_TYPE","DO_CodeTaxe1",
                                    "DO_Taxe1","DO_Type","DO_Imprim","DO_Reliquat",
                                    "AR_Ref","DL_Design","DL_Qte","DL_QteLivre","DL_QteRestant","QteLivre","Total_Qte"
                                };

                                foreach (DataTable dt in new[] { dtAPA, dtAutre })
                                {
                                    foreach (string col in colsCachees)
                                        if (dt.Columns.Contains(col))
                                            dt.Columns[col].ColumnMapping = MappingType.Hidden;

                                    if (dt.Columns.Contains("DO_Piece"))
                                        dt.Columns["DO_Piece"].Caption = "Numéro pièce";
                                    if (dt.Columns.Contains("CT_Intitule"))
                                        dt.Columns["CT_Intitule"].Caption = "Fournisseur";
                                    if (dt.Columns.Contains("TotalTTC"))
                                        dt.Columns["TotalTTC"].Caption = "Montant TTC";
                                    if (dt.Columns.Contains("DL_PrixRU"))
                                        dt.Columns["DL_PrixRU"].Caption = "Prix unitaire";
                                }

                                // ── 5. Binding des GridControls ───────────────────────

                                // gc → APA/BC  (grille simple)
                                BindingSource bsAPA = new BindingSource { DataSource = dtAPA };
                                gc.DataSource = bsAPA;

                                // gc1 → AFA  (grille simple)
                                BindingSource bsAFA = new BindingSource { DataSource = dtAutre };
                                gc1.DataSource = bsAFA;
                            }
                        }
                    }
                    catch (Exception ex) { }
                }

            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                // MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", ...);
            }
        }

        // ════════════════════════════════════════════════════════════════
        //  HELPER : crée un DataSet avec deux tables liées
        // ════════════════════════════════════════════════════════════════
        private static DataSet CreateMasterDetailDataSet(
            string headerTableName,
            string detailTableName,
            bool withAction = false)
        {
            DataSet ds = new DataSet();

            // Table ENTÊTE
            DataTable dtHeader = new DataTable(headerTableName);
            dtHeader.Columns.Add("DO_Piece", typeof(string));
            dtHeader.Columns.Add("CT_Intitule", typeof(string));
            dtHeader.Columns.Add("DL_PrixRU", typeof(string));
            dtHeader.Columns.Add("TotalTTC", typeof(string));
            if (withAction)
                dtHeader.Columns.Add("Action", typeof(string));

            // Table DÉTAIL
            DataTable dtDetail = new DataTable(detailTableName);
            dtDetail.Columns.Add("DO_Piece", typeof(string)); // clé de jointure
            dtDetail.Columns.Add("AR_Ref", typeof(string));
            dtDetail.Columns.Add("DL_Design", typeof(string));
            dtDetail.Columns.Add("DL_Qte", typeof(string));
            dtDetail.Columns.Add("DL_QteLivre", typeof(string));
            dtDetail.Columns.Add("DL_QteRestant", typeof(string));

            ds.Tables.Add(dtHeader);
            ds.Tables.Add(dtDetail);

            // Relation Master → Detail
            string relationName = headerTableName + "To" + detailTableName;
            ds.Relations.Add(
                relationName,
                dtHeader.Columns["DO_Piece"],
                dtDetail.Columns["DO_Piece"]);

            return ds;
        }

        // ════════════════════════════════════════════════════════════════
        //  HELPER : ajoute une ligne dans le DataSet (entête + détail)
        // ════════════════════════════════════════════════════════════════
        private static void AddRowToDataSet(
            DataSet ds,
            string headerTableName,
            string detailTableName,
            DataRow sourceRow,
            string action = null)
        {
            string doPiece = sourceRow["DO_Piece"]?.ToString() ?? "";

            DataTable dtHeader = ds.Tables[headerTableName];
            DataTable dtDetail = ds.Tables[detailTableName];

            // Entête : une seule ligne par DO_Piece
            bool headerExists = dtHeader.AsEnumerable()
                                        .Any(r => r["DO_Piece"].ToString() == doPiece);
            if (!headerExists)
            {
                DataRow headerRow = dtHeader.NewRow();
                headerRow["DO_Piece"] = doPiece;
                headerRow["CT_Intitule"] = sourceRow["CT_Intitule"];
                headerRow["DL_PrixRU"] = sourceRow["DL_PrixRU"];
                headerRow["TotalTTC"] = sourceRow["TotalTTC"];

                if (action != null && dtHeader.Columns.Contains("Action"))
                    headerRow["Action"] = action;

                dtHeader.Rows.Add(headerRow);
            }

            // Détail : une ligne par article
            DataRow detailRow = dtDetail.NewRow();
            detailRow["DO_Piece"] = doPiece;
            detailRow["AR_Ref"] = sourceRow["AR_Ref"];
            detailRow["DL_Design"] = sourceRow["DL_Design"];
            detailRow["DL_Qte"] = sourceRow["DL_Qte"];
            detailRow["DL_QteLivre"] = sourceRow["DL_QteLivre"];
            detailRow["DL_QteRestant"] = sourceRow["DL_QteRestant"];
            dtDetail.Rows.Add(detailRow);
        }

        // ════════════════════════════════════════════════════════════════
        //  HELPER : binding Master-Detail sur un GridControl
        // ════════════════════════════════════════════════════════════════
        private static void BindMasterDetailGrid(
    GridControl gc,
    DataSet ds,
    string headerTableName,
    string detailTableName,
    string relationName,
    bool isLivre)
        {
            try
            {
                // RESET GRID
                gc.DataSource = null;
                gc.LevelTree.Nodes.Clear();
                gc.RepositoryItems.Clear();

                // DATASET
                gc.DataSource = ds;
                gc.DataMember = headerTableName;

                // MASTER VIEW
                GridView masterView = new GridView(gc);

                gc.MainView = masterView;
                gc.ViewCollection.Add(masterView);

                masterView.PopulateColumns(ds.Tables[headerTableName]);

                // OPTIONS
                masterView.OptionsBehavior.Editable = true;
                masterView.OptionsDetail.EnableMasterViewMode = true;
                masterView.OptionsDetail.ShowDetailTabs = false;
                masterView.OptionsDetail.SmartDetailExpand = false;

                masterView.OptionsView.ShowGroupPanel = false;
                masterView.OptionsView.ShowIndicator = false;
                masterView.OptionsView.EnableAppearanceEvenRow = true;
                masterView.OptionsView.ColumnAutoWidth = false;

                masterView.RowHeight = 28;

                // =========================
                // DO_PIECE
                // =========================
                if (masterView.Columns["DO_Piece"] != null)
                {
                    masterView.Columns["DO_Piece"].Caption = "Numéro pièce";
                    masterView.Columns["DO_Piece"].Width = 140;
                }

                // =========================
                // CT_INTITULE
                // =========================
                if (masterView.Columns["CT_Intitule"] != null)
                {
                    masterView.Columns["CT_Intitule"].Caption = "Fournisseur";
                    masterView.Columns["CT_Intitule"].Width = 250;
                }

                // =========================
                // DL_PRIXRU (caché seulement Livre)
                // =========================
                if (masterView.Columns["DL_PrixRU"] != null)
                {
                    masterView.Columns["DL_PrixRU"].Caption = "Prix unitaire";
                    masterView.Columns["DL_PrixRU"].Width = 120;

                    if (isLivre && masterView.Columns["Action"] != null)
                        masterView.Columns["DL_PrixRU"].Visible = false;
                        
                }

                // =========================
                // TOTAL TTC (caché seulement Livre)
                // =========================
                if (masterView.Columns["TotalTTC"] != null)
                {
                    masterView.Columns["TotalTTC"].Caption = "Montant TTC";
                    masterView.Columns["TotalTTC"].Width = 140;

                    if (isLivre)
                        masterView.Columns["TotalTTC"].Visible = false;
                }

                // =========================
                // ACTION (UNIQUEMENT LIVRE)
                // =========================
                if (isLivre && masterView.Columns["Action"] != null)
                {
                    RepositoryItemButtonEdit btnAction = new RepositoryItemButtonEdit();

                    btnAction.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
                    btnAction.AutoHeight = false;
                    btnAction.Buttons.Clear();

                    // 🔵 OUVRIR
                    btnAction.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton()
                    {
                        Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph,
                        Caption = "Ouvrir",
                        Appearance =
                        {
                            BackColor = Color.FromArgb(192, 57, 43),
                            ForeColor = Color.Green,
                            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                            Options = { UseBackColor = true, UseForeColor = true, UseFont = true }
                        }
                    });

                    // 🔴 CLOTURER
                    btnAction.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton()
                    {
                        Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph,
                        Caption = "Clôturer",
                        Appearance =
                        {
                            BackColor = Color.FromArgb(192, 57, 43),
                            ForeColor = Color.Red,
                            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                            Options = { UseBackColor = true, UseForeColor = true, UseFont = true }
                        }
                    });

                    gc.RepositoryItems.Add(btnAction);

                    GridColumn colAction = masterView.Columns["Action"];


                    colAction.ColumnEdit = btnAction;
                    colAction.Caption = "Actions";
                    colAction.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    colAction.AppearanceHeader.Options.UseTextOptions = true;
                    colAction.Width = 120;

                    // 🔥 IMPORTANT : toujours à droite proprement
                    colAction.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Right;

                    colAction.OptionsColumn.AllowEdit = true;
                    colAction.OptionsColumn.ReadOnly = false;
                    colAction.OptionsColumn.AllowFocus = true;

                    colAction.AppearanceCell.TextOptions.HAlignment =
                        DevExpress.Utils.HorzAlignment.Center;

                    colAction.AppearanceCell.BackColor =
                        Color.FromArgb(52, 152, 219);

                    colAction.AppearanceCell.ForeColor = Color.Black;

                    colAction.AppearanceCell.Font =
                        new Font("Segoe UI", 8.5f, FontStyle.Bold);

                    colAction.AppearanceCell.Options.UseBackColor = true;
                    colAction.AppearanceCell.Options.UseForeColor = true;
                    colAction.AppearanceCell.Options.UseFont = true;

                    // CLICK BUTTON

                    btnAction.ButtonClick += (s, e) =>
                    {
                        try
                        {
                            var view = masterView;
                            int rowHandle = view.FocusedRowHandle;
                            if (rowHandle < 0) return;

                            string doPiece = view.GetRowCellValue(rowHandle, "DO_Piece")?.ToString();
                            if (string.IsNullOrEmpty(doPiece)) return;

                            // 🔵 OUVRIR
                            if (e.Button.Index == 0)
                            {
                                ucDocuments.Instance.OuvrirPiece(view, rowHandle);
                            }

                            // 🔴 CLOTURER
                            else if (e.Button.Index == 1)
                            {
                                

                                bool tester1(string cond)
                                {
                                    bool b_test = false;

                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();

                                        string sql = @"
                                        SELECT DL_Qte, QteLivre
                                        FROM F_DOCLIGNE
                                        WHERE DO_Piece LIKE '%' + @DO_Piece + '%'";

                                        using (SqlCommand cmd = new SqlCommand(sql, cn))
                                        {
                                            cmd.Parameters.AddWithValue("@DO_Piece", cond);

                                            using (SqlDataReader reader = cmd.ExecuteReader())
                                            {
                                                while (reader.Read())
                                                {
                                                    var dlQteValue = reader["DL_Qte"];
                                                    var qteLivreValue = reader["QteLivre"];

                                                    Console.WriteLine($"DL_Qte={dlQteValue} | QteLivre={qteLivreValue}");

                                                    decimal dlQte;
                                                    decimal qteLivre;

                                                    if (!decimal.TryParse(dlQteValue?.ToString(), out dlQte))
                                                        continue;

                                                    if (!decimal.TryParse(qteLivreValue?.ToString(), out qteLivre))
                                                        continue;

                                                    if (qteLivre != dlQte)
                                                    {
                                                        b_test = true;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    return b_test;
                                }

                                if (tester1(doPiece))
                                {
                                    MessageBox.Show($"Vous ne pouvez pas clôturer ce document, il y a encore des quantités non livrées!!!!", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                else
                                {
                                    DialogResult result = XtraMessageBox.Show(
                                       $"Clôturer le document {doPiece} ?",
                                       "Confirmation",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Question);

                                    if (result == DialogResult.Yes)
                                    {
                                        AppDbContext _context = new AppDbContext();
                                        var doc = _context.F_DOCENTETE
                                            .FirstOrDefault(d => d.DO_Piece == doPiece);

                                        if (doc == null)
                                        {
                                            MessageBox.Show($"Document introuvable : {doPiece}");
                                            return;
                                        }

                                        doc.DO_Cloture = 1;
                                        _context.SaveChanges();

                                        using (SqlConnection conn = new SqlConnection(connectionString))
                                        {

                                            conn.Open();
                                            string queryss = @"
                                                DISABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                                                UPDATE dbo.F_DOCLIGNE SET
                                                             DO_Piece = @DocPieces
                                                WHERE Do_Piece = @DocPiece AND DL_Qte=QteLivre;
                                                ENABLE TRIGGER ALL ON dbo.F_DOCLIGNE;
                        ";

                                            using (SqlCommand cmds1 = new SqlCommand(queryss, conn))
                                            {
                                                cmds1.Parameters.AddWithValue("@DocPiece", doPiece + '_');
                                                cmds1.Parameters.AddWithValue("@DocPieces", doPiece.Replace("_", ""));
                                                cmds1.ExecuteNonQuery();
                                            }
                                        }
                                        ucDocuments.Instance.RafraichirDonnees();
                                    }
                                }

                               
                            }
                        }
                        catch (Exception ex)
                        {
                            XtraMessageBox.Show(ex.Message);
                        }
                    };
                }

                // =========================
                // DETAIL VIEW
                // =========================
                GridView detailView = new GridView(gc);

                gc.ViewCollection.Add(detailView);

                GridLevelNode levelNode = new GridLevelNode
                {
                    RelationName = relationName,
                    LevelTemplate = detailView
                };

                gc.LevelTree.Nodes.Add(levelNode);

                detailView.PopulateColumns(ds.Tables[detailTableName]);
                detailView.OptionsBehavior.Editable = false;
                detailView.OptionsView.ShowGroupPanel = false;
                detailView.OptionsView.ShowIndicator = false;
                detailView.OptionsView.ColumnAutoWidth = false;

                detailView.Appearance.HeaderPanel.ForeColor = Color.Magenta;
                detailView.Appearance.HeaderPanel.Font = new Font("Segoe UI", 7f, FontStyle.Bold);
                detailView.Appearance.HeaderPanel.Options.UseBackColor = true;
                detailView.Appearance.HeaderPanel.Options.UseForeColor = true;
                detailView.Appearance.HeaderPanel.Options.UseFont = true;

                detailView.Appearance.Row.Font = new Font("Segoe UI", 7f, FontStyle.Regular);

                detailView.RowHeight = 24;

                if (detailView.Columns["DO_Piece"] != null)
                    detailView.Columns["DO_Piece"].Visible = false;

                EnsureDetailColumn(detailView, "AR_Ref", "Référence", 0, 140);
                EnsureDetailColumn(detailView, "DL_Design", "Désignation", 1, 320);
                EnsureDetailColumn(detailView, "DL_Qte", "Qté commandée", 2, 100);
                EnsureDetailColumn(detailView, "DL_QteLivre", "Qté Livrée", 3, 100);
                EnsureDetailColumn(detailView, "DL_QteRestant", "Reste à livrer", 4, 100);

                if (detailView.Columns["DL_Qte"] != null)
                {
                    detailView.Columns["DL_Qte"].AppearanceCell.TextOptions.HAlignment =
                        DevExpress.Utils.HorzAlignment.Far;

                    detailView.Columns["DL_Qte"].DisplayFormat.FormatType =
                        DevExpress.Utils.FormatType.Numeric;

                    detailView.Columns["DL_Qte"].DisplayFormat.FormatString = "N2";
                }

                if (detailView.Columns["DL_QteLivre"] != null)
                {
                    detailView.Columns["DL_QteLivre"].AppearanceCell.TextOptions.HAlignment =
                        DevExpress.Utils.HorzAlignment.Far;

                    detailView.Columns["DL_QteLivre"].DisplayFormat.FormatType =
                        DevExpress.Utils.FormatType.Numeric;

                    detailView.Columns["DL_QteLivre"].DisplayFormat.FormatString = "N2";
                }

                if (detailView.Columns["DL_QteRestant"] != null)
                {
                    detailView.Columns["DL_QteRestant"].AppearanceCell.TextOptions.HAlignment =
                        DevExpress.Utils.HorzAlignment.Far;

                    detailView.Columns["DL_QteRestant"].DisplayFormat.FormatType =
                        DevExpress.Utils.FormatType.Numeric;

                    detailView.Columns["DL_QteRestant"].DisplayFormat.FormatString = "N2";
                }

                detailView.RowStyle += (sender, e) =>
                {
                    e.Appearance.ForeColor = Color.FromArgb(44, 62, 80);
                    e.HighPriority = true;
                };
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Erreur");
            }
        }

        // ════════════════════════════════════════════════════════════════
        //  HELPER : s'assurer qu'une colonne détail existe et est configurée
        // ════════════════════════════════════════════════════════════════
        private static void EnsureDetailColumn(
            GridView view,
            string fieldName,
            string caption,
            int visibleIndex,
            int width)
        {
            GridColumn col = view.Columns[fieldName]
                          ?? view.Columns.AddField(fieldName);

            col.Caption = caption;
            col.VisibleIndex = visibleIndex;
            col.Width = width;
            col.OptionsColumn.ReadOnly = true;
        }

        // ════════════════════════════════════════════════════════════════
        //  GetAllDepots
        // ════════════════════════════════════════════════════════════════
        public static List<F_DEPOT> GetAllDepots()
        {
            List<F_DEPOT> depots = new List<F_DEPOT>();
            string query = "SELECT DE_No, DE_Intitule FROM F_DEPOT";

            using (SqlConnection conn = new SqlConnection(connectionStrings))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            depots.Add(new F_DEPOT
                            {
                                DE_No = (int)reader["DE_No"],
                                DE_Intitule = reader["DE_Intitule"].ToString()
                            });
                    }
                }
                catch (Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}",
                                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);*/
                }
            }
            return depots;
        }

        // ════════════════════════════════════════════════════════════════
        //  GetAllUnites
        // ════════════════════════════════════════════════════════════════
        public static List<P_UNITE> GetAllUnites()
        {
            List<P_UNITE> unites = new List<P_UNITE>();
            string query = "SELECT cbIndice, U_Intitule FROM P_UNITE WHERE U_Intitule <> ''";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            unites.Add(new P_UNITE
                            {
                                cbIndice = (short)reader["cbIndice"],
                                U_Intitule = reader["U_Intitule"].ToString()
                            });
                    }
                }
                catch (Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}",
                                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);*/
                }
            }
            return unites;
        }

        // ════════════════════════════════════════════════════════════════
        //  GetAllDevise
        // ════════════════════════════════════════════════════════════════
        public static List<P_DEVISE> GetAllDevise()
        {
            List<P_DEVISE> devises = new List<P_DEVISE>();
            string query = "SELECT cbMarq, D_Intitule FROM P_DEVISE WHERE D_Intitule <> ''";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            devises.Add(new P_DEVISE
                            {
                                cbMarq = reader["cbMarq"] != DBNull.Value
                                                ? Convert.ToInt32(reader["cbMarq"]) : 0,
                                D_Intitule = reader["D_Intitule"]?.ToString() ?? ""
                            });
                    }
                }
                catch (Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}",
                                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);*/
                }
            }
            return devises;
        }

        // ════════════════════════════════════════════════════════════════
        //  GetAllExpedition
        // ════════════════════════════════════════════════════════════════
        public static List<P_EXPEDITION> GetAllExpedition()
        {
            List<P_EXPEDITION> expeditions = new List<P_EXPEDITION>();
            string query = "SELECT cbMarq, E_Intitule FROM P_EXPEDITION WHERE E_Intitule <> ''";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            expeditions.Add(new P_EXPEDITION
                            {
                                cbMarq = reader["cbMarq"] != DBNull.Value
                                                ? Convert.ToInt32(reader["cbMarq"]) : 0,
                                E_Intitule = reader["E_Intitule"].ToString()
                            });
                    }
                }
                catch (Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}",
                                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);*/
                }
            }
            return expeditions;
        }

        // ════════════════════════════════════════════════════════════════
        //  FiltrerFournisseurs
        // ════════════════════════════════════════════════════════════════
        public static void FiltrerFournisseurs(
            GridControl gc, bool actif, bool sommeil, BindingSource bsFrns)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query;

                    if (actif && !sommeil)
                        query = "SELECT * FROM dbo.F_COMPTET WHERE CT_Sommeil = 0 AND CT_Type = 1 ORDER BY CT_Num";
                    else if (!actif && sommeil)
                        query = "SELECT * FROM dbo.F_COMPTET WHERE CT_Sommeil = 1 AND CT_Type = 1 ORDER BY CT_Num";
                    else if (actif && sommeil)
                        query = "SELECT * FROM dbo.F_COMPTET WHERE CT_Type = 1 ORDER BY CT_Num";
                    else
                    {
                        gc.DataSource = null;
                        return;
                    }

                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        connection.Open();
                        DataTable dt = new DataTable();
                        new SqlDataAdapter(cmd).Fill(dt);
                        rownum = dt.Rows.Count;
                        bsFrns.DataSource = dt;
                        gc.DataSource = bsFrns;
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}",
                                "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);*/
            }
        }

        // ════════════════════════════════════════════════════════════════
        //  GetAllCompteG
        // ════════════════════════════════════════════════════════════════
        public static List<F_COMPTEG> GetAllCompteG()
        {
            List<F_COMPTEG> comptetG = new List<F_COMPTEG>();
            string query = "SELECT DISTINCT CG_Num, CG_Intitule FROM F_COMPTEG WHERE CG_Tiers = 1 AND CG_Sommeil = 0";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                            comptetG.Add(new F_COMPTEG
                            {
                                CG_Num = reader["CG_Num"].ToString(),
                                CG_Intitule = reader["CG_Intitule"].ToString()
                            });
                    }
                }
                catch (Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}",
                                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);*/
                }
            }
            return comptetG;
        }
        
    }
}