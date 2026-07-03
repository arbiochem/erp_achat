using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Models;
using arbioApp.Repositories.ModelsRepository;
using DevExpress.ChartRangeControlClient.Core;
using DevExpress.Charts.Native;
using DevExpress.DashboardCommon.Viewer;
using DevExpress.Utils;
using DevExpress.Xpo;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Alerter;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraExport.Helpers;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraRichEdit.Import.Doc;
using DevExpress.XtraTreeList;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BindingSource = System.Windows.Forms.BindingSource;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class ucDocuments : DevExpress.XtraEditors.XtraUserControl
    {
        private static ucDocuments _instance;
        private System.Data.DataTable dataTable;
        private SqlDataAdapter dataAdapter;
        private SqlConnection connection;
        public static decimal doCours;

        private static string DbPrincipale => ucDocuments.dbNamePrincipale;
        private static string ServerIpPrincipale => ucDocuments.serverIpPrincipale;


        private static string connectionString = $"Server={ServerIpPrincipale};Database={DbPrincipale};" +
                                             $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                             $"Connection Timeout=240;";

        public static ucDocuments Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucDocuments();
                return _instance;
            }
        }

        public ucDocuments()
        {
            ucDocuments.dbNamePrincipale = "TRANSIT";
            ucDocuments.serverNamePrincipale = "SRV-ARB";
            ucDocuments.serverIpPrincipale = "26.53.123.231";
            InitializeComponent();
            CreateDatabaseMenu();
            ChargerDonneesDepuisBDD();

            
            GridView gvDetailLivre = new GridView(gcLivre);
            gvDetailLivre.Name = "gvDetailLivre";

            // Colonnes détail
            gvDetailLivre.Columns.AddField("AR_Ref").VisibleIndex = 0;
            gvDetailLivre.Columns["AR_Ref"].Caption = "Référence";

            gvDetailLivre.Columns.AddField("DL_Design").VisibleIndex = 1;
            gvDetailLivre.Columns["DL_Design"].Caption = "Désignation";

            gvDetailLivre.Columns.AddField("DL_Qte").VisibleIndex = 2;
            gvDetailLivre.Columns["DL_Qte"].Caption = "Qté Livrée";
            gvDetailLivre.Columns["DL_Qte"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvDetailLivre.Columns["DL_Qte"].DisplayFormat.FormatString = "N2";

            // Lier au gcLivre
            gcLivre.LevelTree.Nodes.Add("detail", gvDetailLivre);

            // Activer le master-detail
            gvLivre.OptionsDetail.EnableMasterViewMode = true;
            gvLivre.OptionsDetail.ShowDetailTabs = false;

            // Charger les données au clic sur "+"
            gvLivre.MasterRowExpanded += (s, e) =>
            {
                var masterView = s as GridView;
                var detailView = masterView.GetDetailView(e.RowHandle, 0) as GridView;
                if (detailView == null) return;

                string doPiece = masterView.GetRowCellValue(
                    e.RowHandle, "DO_Piece")?.ToString();
            };

            var dbContext = new AppDbContext();
            _collaborateurRepository = new F_COLLABORATEURRepository(dbContext);
        }

        bool tester1(string cond)
        {
            bool b_test = false;

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();

                string sql = @"
                SELECT DL_Qte, QteLivre
                FROM F_DOCLIGNE
                WHERE DO_Piece = @DO_Piece";

                using (SqlCommand cmd = new SqlCommand(sql, cn))
                {
                    cmd.Parameters.AddWithValue("@DO_Piece", cond);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            decimal dlQte = Convert.ToDecimal(reader["DL_Qte"]);
                            decimal qteLivre = Convert.ToDecimal(reader["QteLivre"]);

                            if (qteLivre == dlQte)
                            {
                                b_test = true;
                                break; // retirez cette ligne si vous voulez parcourir toutes les lignes
                            }
                        }
                    }
                }
            }

            return b_test;
        }
        private void GridViewLivre_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e.Column.FieldName != "Action") return;

            GridView view = sender as GridView;
            DataRow row = view.GetDataRow(e.RowHandle);
            if (row == null) return;

            string doPiece = row["DO_Piece"]?.ToString() ?? "";

            bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "VIEW");

            if (autorise)
            {
                if (tester1(doPiece))
                {
                    if (MessageBox.Show(
                            $"Clôturer le document {doPiece} ?",
                            "Confirmation",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question) != DialogResult.Yes)
                        return;

                    CloturerDocument(doPiece);
                }
                else
                {
                    MessageBox.Show($"Vous ne pouvez pas clôturer ce document, il y a encore des quantités non livrées!!!!", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(
                            "Vous n'avez pas l'autorisation de clôturer un document !",
                            "Transformation bloquée",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
            }
        }

        private void CloturerDocument(string doPiece)
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
                    cmds1.Parameters.AddWithValue("@DocPiece",  doPiece+ '_');
                    cmds1.Parameters.AddWithValue("@DocPieces", doPiece.Replace("_", ""));
                    cmds1.ExecuteNonQuery();
                }
            }

            ChargerDonneesDepuisBDD();
        }

        private void btnNouveauDoc_Click(object sender, EventArgs e)
        {
            frmMenuAchat _frmMenuAchat = new frmMenuAchat();
            _frmMenuAchat.ShowDialog();
        }

        private void ucDocuments_Load(object sender, EventArgs e)
        {
            ChargerDonneesDepuisBDD();
        }

        List<F_DOCENTETE> dotype = new List<F_DOCENTETE>
        {
            new F_DOCENTETE { DO_Type = 10 },
            new F_DOCENTETE { DO_Type = 11 },
            new F_DOCENTETE { DO_Type = 12 },
            new F_DOCENTETE { DO_Type = 13 },
            new F_DOCENTETE { DO_Type = 14 },
            new F_DOCENTETE { DO_Type = 15 },
            new F_DOCENTETE { DO_Type = 16 },
            new F_DOCENTETE { DO_Type = 17 },
            new F_DOCENTETE { DO_Type = 18 },
            new F_DOCENTETE { DO_Type = 200 },
        };

        public int DoTypeSelected;

        public void RafraichirDonnees()
        {
            if (dbNamePrincipale == string.Empty) return;
            ChargerDonneesDepuisBDD();
        }

        public BindingSource BindingEntetes = new BindingSource();

        public void ChargerDonneesDepuisBDD()
        {
            if (BindingEntetes == null)
                BindingEntetes = new BindingSource();
            try
            {
                Entetes.AfficherEntetes_achat(gcEntetes, gcFactures, gcLivre, gcCloture, DoTypeSelected, BindingEntetes);
                gvEntete.BestFitColumns();



                // 4. Lier au GridControl
               
                GridView[] listgv = { gvEntete, gvLivre, gvFacture, gvCloture };

                foreach (var gv in listgv)
                {
                    var colTTC = gv.Columns["DO_TotalTTC"];
                    if (colTTC != null)
                    {
                        colTTC.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                        colTTC.DisplayFormat.FormatString = "N2";
                    }

                    GridColumn coldopiece = null;
                    foreach (GridColumn col in gv.Columns)
                    {
                        if (col.FieldName.ToUpper() == "DO_PIECE")
                        {
                            coldopiece = col;
                            break;
                        }
                    }

                 
                    if (coldopiece != null)
                    {
                        RepositoryItemHyperLinkEdit hyperlink = new RepositoryItemHyperLinkEdit();
                        hyperlink.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
                        gv.GridControl.RepositoryItems.Add(hyperlink);
                        coldopiece.ColumnEdit = hyperlink;
                    }

                    gv.RowCellClick -= Gv_RowCellClick;
                    gv.RowCellClick += Gv_RowCellClick;

                    gvLivre.RowCellClick -= Gv_RowCellClick;
                    gvLivre.RowCellClick += Gv_RowCellClick;


                    var colDLQte = gv.Columns["DL_Qte"]; 
                    if (colDLQte != null)
                    {
                        colDLQte.Visible = (gv == gvLivre);
                        colDLQte.Caption = "Qté livrée";
                    }

                    var colDLRef = gv.Columns["AR_Ref"];
                    if (colDLRef != null)
                    {
                        colDLRef.Visible = (gv == gvLivre);
                        colDLRef.Caption = "Référence";
                    }

                    var colDLDesign = gv.Columns["DL_Design"];
                    if (colDLDesign != null)
                    {
                        colDLDesign.Visible = (gv == gvLivre);
                        colDLDesign.Caption = "Designation";
                    }


                    var colDLPrixRU = gv.Columns["DL_PrixRU"];
                    if (colDLPrixRU != null)
                    {
                        colDLPrixRU.Visible = (gv == gvCloture);
                        colDLPrixRU.Caption = "Prix de revient";
                    }


                    var cols = gvLivre.Columns["TotalTTC"];
                    if (cols != null)
                        cols.Visible = false;

                    var colss = gvCloture.Columns["TotalTTC"];
                    if (colss != null)
                        colss.Visible = false;
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                /*MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);*/
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dbNamePrincipale == string.Empty) return;
        }

        public void RafraichirListeDocumentsParDoType(int doType)
        {
            // listBox1.SelectedValue = doType;
        }

        private void gvEntete_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
        }

        private void gvEntete_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName == "DO_Imprim" && e.Value != null)
            {
                if (int.TryParse(e.Value.ToString(), out _))
                    e.DisplayText = "";
            }
        }

        private void gvEntete_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                if (e.Column.FieldName == "DO_Imprim")
                {
                    int value = Convert.ToInt32(gvEntete.GetRowCellValue(e.RowHandle, "DO_Imprim"));
                    e.Appearance.Image = value == 1 ? imageCollection1.Images[0] : imageCollection1.Images[1];
                }
            }
            catch (Exception) { }
        }

        private void gvEntete_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                if (e.Column.FieldName == "DO_Imprim")
                {
                    int value = Convert.ToInt32(gvEntete.GetRowCellValue(e.RowHandle, "DO_Imprim"));
                    Image icon = value == 1 ? imageCollection1.Images[0] : imageCollection1.Images[1];

                    int iconWidth = 16;
                    int iconHeight = 16;
                    int x = e.Bounds.X + (e.Bounds.Width - iconWidth) / 2;
                    int y = e.Bounds.Y + (e.Bounds.Height - iconHeight) / 2;

                    e.Graphics.DrawImage(icon, x, y, iconWidth, iconHeight);
                    e.Handled = true;
                }
            }
            catch (Exception) { }
        }

        // ── Champs statiques partagés ────────────────────────────────────────
        public static DateTime doDate;
        public static DateTime doDateLivrPrev;
        public static int doImprim;
        public static string doTiers;
        public static int doStatut;
        public static int doReliquat;
        public static int deno;
        public static string doRef;
        public static int TypeAchat;
        public static int doExpedit;
        public static string doPiece;
        public static int doTaxe1;
        public static string doCodeTaxe1;
        public static int? CoNo;
        public static string doEntete;
        public static string a_type;

        private F_COLLABORATEURRepository _collaborateurRepository;
        F_COLLABORATEUR doCollaborateur;

        public static string dbNamePrincipale = string.Empty;
        public static string serverNamePrincipale = string.Empty;
        public static string serverIpPrincipale = string.Empty;

        // ── Ouverture d'un document ──────────────────────────────────────────
        private void Gv_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            GridView gv = sender as GridView;
            if (gv == null) return;
            if (e.Column == null) return;

            // Pour tous les autres GridView → colonne DO_Piece
            if (e.Column.FieldName == "DO_Piece")
            {
                object doPieceObj = gv.GetRowCellValue(e.RowHandle, "DO_Piece");

                bool autorise = frmMenuAchat.verifier_droit("Facture", "VIEW");
                bool autorise1 = frmMenuAchat.verifier_droit("Projet d'achat", "VIEW");
                bool autorise2 = frmMenuAchat.verifier_droit("Bon de commande", "VIEW");
                bool autorise3 = frmMenuAchat.verifier_droit("Facture", "SAISIE_QTE_LIVRE");

                if (doPieceObj.ToString().StartsWith("APA"))
                {
                    if (autorise1)
                    {
                        OuvrirPiece(gv, e.RowHandle);
                        return;
                    }
                    else
                    {
                        MessageBox.Show(
                                   "Vous n'avez pas l'autorisation de modifier un projet d'achat !",
                                   "Transformation bloquée",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Error
                               );
                    }
                }
                else if (doPieceObj.ToString().StartsWith("ABC"))
                {
                    if (autorise2)
                    {
                        OuvrirPiece(gv, e.RowHandle);
                        return;
                    }
                    else
                    {
                        MessageBox.Show(
                                   "Vous n'avez pas l'autorisation de modifier un bon de commande !",
                                   "Transformation bloquée",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Error
                               );
                    }
                }
                else
                {

                    if (autorise)
                    {
                        OuvrirPiece(gv, e.RowHandle);
                        return;
                    }
                    else
                    {
                        if (autorise3)
                        {
                            OuvrirPiece(gv, e.RowHandle);
                            return;
                        }
                        else
                        {
                            MessageBox.Show(
                                       "Vous n'avez pas l'autorisation de modifier une facture !",
                                       "Transformation bloquée",
                                       MessageBoxButtons.OK,
                                       MessageBoxIcon.Error
                                   );
                        }
                    }
                }
            }

            // Pour gvLivre → colonne Action (bouton Ouvrir)
            if (gv == gvLivre && e.Column.FieldName == "Action")
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "VIEW");

                if (autorise)
                {
                    OuvrirPiece(gv, e.RowHandle);
                    return;
                }
                else
                {
                    MessageBox.Show(
                               "Vous n'avez pas l'autorisation de modifier un bon de réception !",
                               "Transformation bloquée",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error
                           );
                }
            }
        }

        public void OuvrirPiece(GridView gv, int rowHandle)
        {
            try
            {
                if (rowHandle < 0)
                {
                    MessageBox.Show("Veuillez sélectionner une ligne.", "Avertissement",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Récupérer DO_Piece depuis la ligne exacte cliquée
                string dopiece_selected = gv.GetRowCellValue(rowHandle, "DO_Piece")?.ToString()?.Trim();
                if (string.IsNullOrEmpty(dopiece_selected)) return;

                // Charger toutes les données depuis la base avec DO_Piece
                using (var context = new AppDbContext())
                {
                    if (dopiece_selected.StartsWith("AFA"))
                    {
                        var doc_ = context.F_DOCENTETE
                            .FirstOrDefault(d => d.DO_Piece.Trim() == dopiece_selected);

                        if (doc_ == null)
                        {
                            var doc = context.F_DOCENTETE
                            .FirstOrDefault(d => d.DO_Piece.Trim() == dopiece_selected.Replace("AFA", "ABR"));

                            if (doc == null)
                            {
                                MessageBox.Show("Document introuvable.", "Avertissement",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            doTiers = doc.DO_Tiers?.Trim() ?? "";
                            doRef = doc.DO_Ref?.Trim() ?? "";
                            doStatut = (int)(2);
                            doDate = doc.DO_Date ?? DateTime.Now;
                            doDateLivrPrev = doc.DO_DateLivr ?? DateTime.Now;
                            int? CO_No = doc.CO_No ?? 0;
                            CoNo = CO_No;
                            doEntete = doc.DO_Coord01?.Trim() ?? "";
                            deno = (int)(doc.DE_No ?? 0);
                            doCodeTaxe1 = doc.DO_CodeTaxe1?.Trim() ?? "";
                            doTaxe1 = (int)(doc.DO_Taxe1 ?? 0);
                            doExpedit = (int)(doc.DO_Expedit ?? 0);
                            TypeAchat = 16;
                            doPiece = doc.DO_Piece.Replace("ABR","AFA")?.Trim() ?? "";
                            doImprim = (int)(doc.DO_Imprim ?? 0);
                            doReliquat = (int)(doc.DO_Reliquat ?? 0);

                            if (doc.DO_Piece.StartsWith("ABC"))
                            {
                                a_type = "Bon de commande";
                            }
                            else if (doc.DO_Piece.StartsWith("AFA"))
                            {
                                a_type = "Facture";
                            }
                            else if (doc.DO_Piece.StartsWith("ABR"))
                            {
                                a_type = "Bon de livraison";
                            }


                            if (CO_No == 0)
                            {
                                MessageBox.Show("Pas de collaborateur pour ceci", "Avertissement",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            doCollaborateur = _collaborateurRepository.GetBy_CO_No(CO_No);
                            if (doCollaborateur == null)
                            {
                                MessageBox.Show("Collaborateur non trouvé.");
                                return;
                            }

                            // Déplacer BindingEntetes vers la bonne ligne
                            if (BindingEntetes.List != null)
                            {
                                for (int i = 0; i < BindingEntetes.List.Count; i++)
                                {
                                    DataRowView row = BindingEntetes.List[i] as DataRowView;
                                    if (row != null &&
                                        row["DO_Piece"]?.ToString()?.Trim() == dopiece_selected)
                                    {
                                        BindingEntetes.Position = i;
                                        break;
                                    }
                                }
                            }

                            string nodoc = doPiece.Substring(3, 8);
                            string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";
                            if (!Directory.Exists(destinationFolderdoc))
                                Directory.CreateDirectory(destinationFolderdoc);

                            frmEditDocument editForm = new frmEditDocument(doPiece.Replace("AFA", "ABR"), a_type, this, BindingEntetes);
                            editForm.ShowDialog();

                            ChargerDonneesDepuisBDD();
                        }
                        else
                        {
                            var doc = context.F_DOCENTETE
                            .FirstOrDefault(d => d.DO_Piece.Trim() == dopiece_selected);

                            if (doc == null)
                            {
                                MessageBox.Show("Document introuvable.", "Avertissement",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            doTiers = doc.DO_Tiers?.Trim() ?? "";
                            doRef = doc.DO_Ref?.Trim() ?? "";
                            doStatut = (int)(doc.DO_Statut ?? 0);
                            doDate = doc.DO_Date ?? DateTime.Now;
                            doDateLivrPrev = doc.DO_DateLivr ?? DateTime.Now;
                            int? CO_No = doc.CO_No ?? 0;
                            CoNo = CO_No;
                            doEntete = doc.DO_Coord01?.Trim() ?? "";
                            deno = (int)(doc.DE_No ?? 0);
                            doCodeTaxe1 = doc.DO_CodeTaxe1?.Trim() ?? "";
                            doTaxe1 = (int)(doc.DO_Taxe1 ?? 0);
                            doExpedit = (int)(doc.DO_Expedit ?? 0);
                            TypeAchat = doc.DO_Type ?? 0;
                            doPiece = doc.DO_Piece?.Trim() ?? "";
                            doImprim = (int)(doc.DO_Imprim ?? 0);
                            doReliquat = (int)(doc.DO_Reliquat ?? 0);

                            if (dopiece_selected.StartsWith("AFA"))
                            {
                                a_type = "Facture";
                            }
                            else if (dopiece_selected.StartsWith("ABR"))
                            {
                                a_type = "Bon de livraison";
                            }


                            if (CO_No == 0)
                            {
                                MessageBox.Show("Pas de collaborateur pour ceci", "Avertissement",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            doCollaborateur = _collaborateurRepository.GetBy_CO_No(CO_No);
                            if (doCollaborateur == null)
                            {
                                MessageBox.Show("Collaborateur non trouvé.");
                                return;
                            }

                            // Déplacer BindingEntetes vers la bonne ligne
                            if (BindingEntetes.List != null)
                            {
                                for (int i = 0; i < BindingEntetes.List.Count; i++)
                                {
                                    DataRowView row = BindingEntetes.List[i] as DataRowView;
                                    if (row != null &&
                                        row["DO_Piece"]?.ToString()?.Trim() == dopiece_selected)
                                    {
                                        BindingEntetes.Position = i;
                                        break;
                                    }
                                }
                            }

                            string nodoc = doPiece.Substring(3, 8);
                            string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";
                            if (!Directory.Exists(destinationFolderdoc))
                                Directory.CreateDirectory(destinationFolderdoc);

                            frmEditDocument editForm = new frmEditDocument(doPiece, a_type, this, BindingEntetes);
                            editForm.ShowDialog();

                            

                            ChargerDonneesDepuisBDD();
                        }
                    }
                    else
                    {
                        var doc = context.F_DOCENTETE
                            .FirstOrDefault(d => d.DO_Piece.Trim() == dopiece_selected);

                        if (doc == null)
                        {
                            MessageBox.Show("Document introuvable.", "Avertissement",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        doTiers = doc.DO_Tiers?.Trim() ?? "";
                        doRef = doc.DO_Ref?.Trim() ?? "";
                        doStatut = (int)(doc.DO_Statut ?? 0);
                        doDate = doc.DO_Date ?? DateTime.Now;
                        doDateLivrPrev = doc.DO_DateLivr ?? DateTime.Now;
                        int? CO_No = doc.CO_No ?? 0;
                        CoNo = CO_No;
                        doEntete = doc.DO_Coord01?.Trim() ?? "";
                        deno = (int)(doc.DE_No ?? 0);
                        doCodeTaxe1 = doc.DO_CodeTaxe1?.Trim() ?? "";
                        doTaxe1 = (int)(doc.DO_Taxe1 ?? 0);
                        doExpedit = (int)(doc.DO_Expedit ?? 0);
                        TypeAchat = doc.DO_Type ?? 0;
                        doPiece = doc.DO_Piece?.Trim() ?? "";
                        doImprim = (int)(doc.DO_Imprim ?? 0);
                        doReliquat = (int)(doc.DO_Reliquat ?? 0);

                        if (doc.DO_Piece.StartsWith("ABC"))
                        {
                            a_type = "Bon de commande";
                        }
                        else if (doc.DO_Piece.StartsWith("AFA"))
                        {
                            a_type = "Facture";
                        }
                        else if (doc.DO_Piece.StartsWith("ABR"))
                        {
                            a_type = "Bon de livraison";
                        }
                        else if (doc.DO_Piece.StartsWith("APA"))
                        {
                            a_type = "Projet d'achat";
                        }


                        if (CO_No == 0)
                        {
                            MessageBox.Show("Pas de collaborateur pour ceci", "Avertissement",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        doCollaborateur = _collaborateurRepository.GetBy_CO_No(CO_No);
                        if (doCollaborateur == null)
                        {
                            MessageBox.Show("Collaborateur non trouvé.");
                            return;
                        }

                        // Déplacer BindingEntetes vers la bonne ligne
                        if (BindingEntetes.List != null)
                        {
                            for (int i = 0; i < BindingEntetes.List.Count; i++)
                            {
                                DataRowView row = BindingEntetes.List[i] as DataRowView;
                                if (row != null &&
                                    row["DO_Piece"]?.ToString()?.Trim() == dopiece_selected)
                                {
                                    BindingEntetes.Position = i;
                                    break;
                                }
                            }
                        }

                        string nodoc = doPiece.Substring(3, 8);
                        string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";
                        if (!Directory.Exists(destinationFolderdoc))
                            Directory.CreateDirectory(destinationFolderdoc);

                        frmEditDocument editForm = new frmEditDocument(doPiece, a_type, this, BindingEntetes);
                        editForm.ShowDialog();

                        ChargerDonneesDepuisBDD();
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Boutons ──────────────────────────────────────────────────────────
        private void btnOuvrirDoc_Click(object sender, EventArgs e)
        {
            OuvrirPiece(gvEntete, gvEntete.FocusedRowHandle);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            ChargerDonneesDepuisBDD();
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            gcEntetes.ShowPrintPreview();
        }

        // ── Menu ─────────────────────────────────────────────────────────────
        BarSubItem fileMenu1;
        private void CreateDatabaseMenu()
        {
            DataTable databases = GetDatabasesFromF_ACHATFILES();

            BarSubItem fileMenu = new BarSubItem();
            fileMenu.Caption = "Fichier";

            fileMenu1 = new BarSubItem();
            fileMenu1.Caption = "Etat";
            fileMenu1.Enabled = false;

            barManager1.Items.Add(fileMenu);
            barManager1.Items.Add(fileMenu1);
            barManager1.MainMenu.LinksPersistInfo.Add(new LinkPersistInfo(fileMenu));
            barManager1.MainMenu.LinksPersistInfo.Add(new LinkPersistInfo(fileMenu1));

            BarButtonItem dbItem1 = new BarButtonItem();
            dbItem1.Caption = "Etat global";
            dbItem1.ItemClick += DbItem_Item1Click;
            fileMenu1.AddItem(dbItem1);

            BarSubItem autItem = new BarSubItem();
            autItem.Caption = "Autorisation";
            fileMenu.ItemLinks.Add(autItem);

            BarButtonItem genItem = new BarButtonItem();
            genItem.Caption = "Générer Projet d'achat";
            genItem.ItemClick += genItem_Itemclick;
            fileMenu.ItemLinks.Add(genItem);

            BarButtonItem autGlogale = new BarButtonItem();
            autGlogale.Caption = "Globale";
            autGlogale.ItemClick += autGlogale_ItemClick;

            BarButtonItem autAchat = new BarButtonItem();
            autAchat.Caption = "Type de documents";
            autAchat.ItemClick += autAchat_ItemClick;

            autItem.ItemLinks.Add(autGlogale);
            autItem.ItemLinks.Add(autAchat).BeginGroup = true;
        }

        private void genItem_Itemclick(object sender, ItemClickEventArgs e)
        {
            bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "VIEW");

            if (autorise)
            {
                frm_generer frmGen = new frm_generer();
                frmGen.ShowDialog();
                ChargerDonneesDepuisBDD();
                return;
            }
            else
            {
                MessageBox.Show(
                           "Vous n'avez pas l'autorisation de créer un projet d'achat !",
                           "Transformation bloquée",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error
                       );
            }
        }

        private void DbItem_Item1Click(object sender, ItemClickEventArgs e)
        {
            frm_Etat_general frmEtat = new frm_Etat_general();
            frmEtat.ShowDialog();
        }

        private void autGlogale_ItemClick(object sender, ItemClickEventArgs e)
        {
            frmAutorisation _frmAutorisation = new frmAutorisation();
            _frmAutorisation.ShowDialog();
        }

        private void autAchat_ItemClick(object sender, ItemClickEventArgs e)
        {
            frmAutorisations_achat _frmAutorisation_achat = new frmAutorisations_achat();
            _frmAutorisation_achat.ShowDialog();
        }

        private void autItem_ItemClick(object sender, ItemClickEventArgs e)
        {
            frmAutorisation _frmAutorisation = new frmAutorisation();
            _frmAutorisation.ShowDialog();
        }

        private DataTable GetDatabasesFromF_ACHATFILES()
        {
            DataTable dt = new DataTable();
            string query = "SELECT ID, DBNAME, SERVERNAME, SERVERIP FROM F_ACHATFILES";
            string cnfiles = $"Server={FrmMdiParent.DataSourceNameValueParent};Database=arbapp;" +
                              $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                              $"Connection Timeout=240;";

            using (SqlConnection conn = new SqlConnection(cnfiles))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                dt.Load(cmd.ExecuteReader());
            }
            return dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "VIEW");

            if (autorise)
            {
                frm_generer frmGen = new frm_generer();
                frmGen.ShowDialog();
                ChargerDonneesDepuisBDD();
                return;
            }
            else
            {
                MessageBox.Show(
                           "Vous n'avez pas l'autorisation de créer un projet d'achat !",
                           "Transformation bloquée",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error
                       );
            }
            
        }

        private void gcLivre_Load(object sender, EventArgs e)
        {
            // 1. Configurer le master
            
        }

        private DataTable ChargerDetailLivre(string doPiece)
        {
            string query = @"
                SELECT 
                    AR_Ref,
                    DL_Design      AS Designation,
                    DL_Qte         AS DL_Qte
                FROM dbo.F_DOCLIGNE
                WHERE DO_Piece = @doPiece
                ORDER BY DL_Ligne ASC";

            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@doPiece", doPiece);
                conn.Open();
                new SqlDataAdapter(cmd).Fill(dt);
            }
            return dt;
        }
    }
}