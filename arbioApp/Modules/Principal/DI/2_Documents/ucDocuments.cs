using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using arbioApp.Modules.Principal.DI._2_Documents;
using System.Data.SqlClient;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.Charts.Native;
using System.Net;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraExport.Helpers;
using DevExpress.DashboardCommon.Viewer;
using DevExpress.Utils;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraTreeList;
using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using DevExpress.ChartRangeControlClient.Core;
using BindingSource = System.Windows.Forms.BindingSource;
using DevExpress.Xpo;
using arbioApp.Modules.Principal.DI.Models;
using DevExpress.XtraBars;
using System.IO;
using DevExpress.XtraCharts.Native;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class ucDocuments : DevExpress.XtraEditors.XtraUserControl
    {
        private static ucDocuments _instance;
        private System.Data.DataTable dataTable;
        private SqlDataAdapter dataAdapter;
        private SqlConnection connection;
        public static string connectionString;
        public static decimal doCours;

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

            var dbContext = new AppDbContext();
            _collaborateurRepository = new F_COLLABORATEURRepository(dbContext);
        }

        private void GridViewLivre_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e.Column.FieldName != "Action") return;

            GridView view = sender as GridView;
            DataRow row = view.GetDataRow(e.RowHandle);
            if (row == null) return;

            string doPiece = row["DO_Piece"]?.ToString() ?? "";

            if (MessageBox.Show(
                    $"Clôturer le document {doPiece} ?",
                    "Confirmation",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            CloturerDocument(doPiece);
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
            try
            {
                Entetes.AfficherEntetes_achat(gcEntetes, gcFactures, gcLivre, gcCloture, DoTypeSelected, BindingEntetes);
                gvEntete.BestFitColumns();

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
            if (e.Column.FieldName != "DO_Piece") return;

            OuvrirPiece(gv, e.RowHandle);
        }

        private void OuvrirPiece(GridView gv, int rowHandle)
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
                    a_type = gvEntete.GetFocusedRowCellValue("A_TYPE")?.ToString();

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
            frm_generer frmGen = new frm_generer();
            frmGen.ShowDialog();
            ChargerDonneesDepuisBDD();
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
            frm_generer frmGen = new frm_generer();
            frmGen.ShowDialog();
            ChargerDonneesDepuisBDD();
        }
    }
}