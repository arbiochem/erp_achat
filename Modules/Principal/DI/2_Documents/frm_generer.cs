using arbioApp.DTO;
using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Services;
using DevExpress.Data;
using DevExpress.DataProcessing;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Layout;
using DevExpress.XtraGrid.Views.Layout.ViewInfo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Humanizer.In;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frm_generer : Form
    {
        private LayoutView layoutView;
        private DataTable dt;
        private Dictionary<int, Rectangle> cardBoundsDict = new Dictionary<int, Rectangle>();
        private readonly F_DOCENTETEService _f_DOCENTETEService;
        private readonly F_DOCENTETERepository _f_DOCENTETERepository;
        private Dictionary<string, bool> checkStates = new Dictionary<string, bool>();
        string _prefix = string.Empty;
        List<F_DOCENTETE> _listeDocs;
        private static string connectionString = "Server=26.53.123.231;Database=ARBIOCHEM;User Id=Dev;Password=1234;";
        private string connectionStrings = "Server=26.53.123.231;Database=TRANSIT;User Id=Dev;Password=1234;";
        private string doPiece;
        static int year = DateTime.Now.Year;
        private string ct_num;
        private string dodo_piece;
        private DateTime dts;
        private DateTime dtc;
        private Dictionary<string, Rectangle> articleLineBounds = new Dictionary<string, Rectangle>();
        private Dictionary<string, bool> checkStatesFournisseurs = new Dictionary<string, bool>();

            private Dictionary<string, Dictionary<string, bool>> checkStatesArticles
            = new Dictionary<string, Dictionary<string, bool>>();
        public frm_generer()
        {
            _prefix = "APA";
            InitializeComponent();
        }

        public static string recuperer_ctnum(string fournisseur,string connStr)
        {
            string query = "SELECT CT_Num FROM F_COMPTET WHERE CT_Intitule LIKE @intitule";

            // Récupérer le dernier numéro
            using (var connection = new SqlConnection(connStr))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.Clear();
                    command.Parameters.Add("@intitule", SqlDbType.VarChar).Value = $"%{fournisseur.Trim()}%";

                    object result = command.ExecuteScalar();
                    string val = (result != null && result != DBNull.Value) ? result.ToString() : "";

                    return val;
                }
            }
        }
        public static string GetNextInvoiceNumber(string prefix)
        {
            string query = "SELECT CurrentNumber FROM ARB_ACHAT_DOPIECE WHERE Prefix = @Prefix AND Year = @Year";

            // Récupérer le dernier numéro
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Prefix", prefix);
                    command.Parameters.AddWithValue("@Year", year);

                    object result = command.ExecuteScalar();
                    int currentNumber = result != DBNull.Value ? (int)result : 0;  // Si pas de résultat, commencer à 0

                    connection.Close();
                    return $"{prefix}{year}{currentNumber:D4}";
                }

            }
        }

        private void frm_generer_Load(object sender, EventArgs e)
        {
            ChargerDonneesDepuisBDD();
        }

        private string GetRowKey(int rowHandle)
        {
            string ctNum = layoutView.GetRowCellValue(rowHandle, "CT_Num")?.ToString() ?? "";
            //string arRef = layoutView.GetRowCellValue(rowHandle, "AR_Ref")?.ToString() ?? "";
            // return $"{ctNum}|{arRef}";
            return $"{ctNum}";
        }

        private Guid getcbCreationUserGuid(string usermail)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(
                    @"SELECT PROT_Guid
                      FROM   F_PROTECTIONCIAL
                      WHERE  PROT_EMail = @usermail", conn))
                {
                    cmd.Parameters.Add("@usermail", SqlDbType.NVarChar, 256).Value = usermail;
                    conn.Open();
                    Guid? result = cmd.ExecuteScalar() as Guid?;
                    return result.HasValue ? result.Value : Guid.Empty;
                }
            }
        }

        string FormatNumber(string value)
        {
            if (double.TryParse(value, out double number))
                return number.ToString("N2", new System.Globalization.CultureInfo("fr-FR"));
            return "0,00";
        }
        private string InsertFDOCENTETE(string ct_Num, string fournisseur, string connStr)
        {
            using (SqlConnection conn = new SqlConnection(connectionStrings))
            {
                conn.Open();
                new SqlCommand("DISABLE TRIGGER ALL ON F_DOCENTETE", conn).ExecuteNonQuery();
                ct_Num = recuperer_ctnum(fournisseur, connStr);


                string sql = @"
                    INSERT INTO F_DOCENTETE
                        (DO_Domaine,DO_Type, DO_Piece, DO_Date,DO_Tiers)
                    VALUES
                        (1,10, @doPiece, @doDate,@dotiers)";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@doPiece", SqlDbType.VarChar, 13).Value = doPiece.PadRight(13).Trim();
                    cmd.Parameters.Add("@dotiers", SqlDbType.Char, 13).Value = ct_Num.PadRight(13);

                    // DO_Date est DateTime dans Sage 100
                    cmd.Parameters.Add("@doDate", SqlDbType.DateTime).Value = DateTime.Today;

                    cmd.ExecuteNonQuery();
                }

                new SqlCommand("ENABLE TRIGGER ALL ON F_DOCENTETE", conn).ExecuteNonQuery();


                string queryUpdate = "UPDATE ARB_ACHAT_DOPIECE SET CurrentNumber = CurrentNumber + 1 WHERE Prefix = @Prefix AND Year = @Year";
                string querySelect = "SELECT CurrentNumber FROM ARB_ACHAT_DOPIECE WHERE Prefix = @Prefix AND Year = @Year";

                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // 1. Incrémenter
                    using (var cmdUpdate = new SqlCommand(queryUpdate, connection))
                    {
                        cmdUpdate.Parameters.Add("@Prefix", SqlDbType.VarChar).Value = _prefix;
                        cmdUpdate.Parameters.Add("@Year", SqlDbType.Int).Value = year;
                        cmdUpdate.ExecuteNonQuery();
                    }

                    // 2. Récupérer la nouvelle valeur
                    using (var cmdSelect = new SqlCommand(querySelect, connection))
                    {
                        cmdSelect.Parameters.Add("@Prefix", SqlDbType.VarChar).Value = _prefix;
                        cmdSelect.Parameters.Add("@Year", SqlDbType.Int).Value = year;

                        object result = cmdSelect.ExecuteScalar();
                        int currentNumber = result != null && result != DBNull.Value ? (int)result : 1;

                        doPiece = $"{_prefix}{year}{currentNumber:D4}";
                        return $"{_prefix}{year}{currentNumber:D4}";
                    }
                }

                return doPiece;
            }
        }

        private void InsertFDOCLIGNE(string doPiece, string ar_Ref, string designation,
                                      double qteACommander = 0,decimal pu=0,decimal montant=0,string ct_num="",int depot=0,decimal poids=0,string connStr="")
        {
            using (SqlConnection conn = new SqlConnection(connectionStrings))
            {
                conn.Open();

                int dlLigne;
                using (SqlCommand cmdLigne = new SqlCommand(
                    @"SELECT ISNULL(MAX(DL_Ligne), 0) + 1
                      FROM   F_DOCLIGNE
                      WHERE  DO_Piece = @doPiece AND DO_Type = 1", conn))
                {
                    cmdLigne.Parameters.AddWithValue("@doPiece", doPiece);
                    dlLigne = Convert.ToInt32(cmdLigne.ExecuteScalar());
                }

                string user=FrmMdiParent._id_user.ToString();

                new SqlCommand("DISABLE TRIGGER ALL ON F_DOCLIGNE", conn).ExecuteNonQuery();

                string sql = @"
                INSERT INTO F_DOCLIGNE
                    (DO_Domaine, DO_Type, DO_Piece, DL_Ligne, AR_Ref, DL_Design, PF_Num, DL_No,DL_Qte,DL_PrixUnitaire,DL_CMUP,DL_PUTTC,DL_MontantHT,DL_MontantTTC,CT_Num,DE_No,cbDE_No,DO_Date,DL_DateBC,DL_DateBL,DL_PoidsNet,cbCreationUser,retenu)
                VALUES
                    (1, 10, @doPiece, @dlLigne, @arRef, @design, @pfNum, @dlNo, @dlqte,@dlpu,@dlpu,@dlpu,@montant,@montant,@ctnum,@deno,@deno,@date,@date,@date,@poids,@user,0)";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    int dlNo = GetNextIDLNo(conn); // ← unique à chaque ligne

                    cmd.Parameters.Add("@doPiece", SqlDbType.VarChar, 13).Value = doPiece.PadRight(13).Trim();
                    cmd.Parameters.Add("@dlqte", SqlDbType.Int, 13).Value = qteACommander;
                    cmd.Parameters.Add("@dlLigne", SqlDbType.Int).Value = dlLigne;
                    cmd.Parameters.Add("@arRef", SqlDbType.Char, 18).Value = (ar_Ref ?? string.Empty).PadRight(18);
                    cmd.Parameters.Add("@design", SqlDbType.Char, 69).Value = (designation ?? string.Empty).PadRight(69);
                    cmd.Parameters.Add("@pfNum", SqlDbType.Char, 8).Value = "".PadRight(8);
                    cmd.Parameters.Add("@dlNo", SqlDbType.Int).Value = dlNo;
                    cmd.Parameters.Add("@dlpu", SqlDbType.Decimal).Value = pu;
                    cmd.Parameters.Add("@montant", SqlDbType.Decimal).Value = montant;
                    cmd.Parameters.Add("@ctnum", SqlDbType.Char).Value = ct_num;
                    cmd.Parameters.Add("@deno", SqlDbType.Int).Value = depot;
                    cmd.Parameters.Add("@date", SqlDbType.DateTime).Value = DateTime.Today;
                    cmd.Parameters.Add("@poids", SqlDbType.Decimal).Value = poids;
                    cmd.Parameters.Add("@user", SqlDbType.Char).Value = user;

                    cmd.ExecuteNonQuery();
                }

                new SqlCommand("ENABLE TRIGGER ALL ON F_DOCLIGNE", conn).ExecuteNonQuery();
            }
        }

        private int GetNextIDLNo(SqlConnection conn)
        {
            string query = "SELECT ISNULL(MAX(DL_No), 0) + 1 FROM F_DOCLIGNE";
            using (var cmd = new SqlCommand(query, conn))
            {
                return (int)cmd.ExecuteScalar();
            }
        }

        string devise;
        public void ChargerDonneesDepuisBDD()
        {
            string connectionString1 = "Server=26.53.123.231;Database=ARBIOCHEM;User Id=Dev;Password=1234;";
            string connectionString2 = "Server=26.53.123.231;Database=ACTIVO;User Id=Dev;Password=1234;";

            // ── Dictionnaires d'état ────────────────────────────────────────────────────
            // checkStates       : clé = "CT_Num|Source|AR_Ref",   valeur = bool coché
            // articleLineBounds : clé = "rowHandle|AR_Ref|Source", valeur = Rectangle ligne
            // cardBoundsDict    : clé = rowHandle,                 valeur = Rectangle carte
            // articleConnectionMap : clé = "CT_Num|Source|AR_Ref", valeur = connectionString

            var checkStates = new Dictionary<string, bool>();
            var articleLineBounds = new Dictionary<string, Rectangle>();
            var cardBoundsDict = new Dictionary<int, Rectangle>();
            var articleConnectionMap = new Dictionary<string, string>();

            DataTable dt = null;
            LayoutView layoutView = null;

            // ══════════════════════════════════════════════════════════════════════════════
            // 1. CHARGEMENT DONNÉES DES 2 BASES
            // ══════════════════════════════════════════════════════════════════════════════

            void LoadData()
            {
                dt = new DataTable();
                dt.Columns.Add("CT_Num", typeof(string));
                dt.Columns.Add("CT_Intitule", typeof(string));
                dt.Columns.Add("StockADate", typeof(double));
                dt.Columns.Add("StockMinimum", typeof(double));
                dt.Columns.Add("StockMaximum", typeof(double));
                dt.Columns.Add("Source", typeof(string));

                string query = @"
        SELECT 
            CT_Num, CT_Intitule,
            SUM(StockADate)   AS StockADate,
            MIN(StockMinimum) AS StockMinimum,
            MAX(StockMaximum) AS StockMaximum
        FROM VW_SEUIL_MIN_STOCK_SUM
        GROUP BY CT_Num, CT_Intitule
        ORDER BY CT_Intitule";

                // ── ARBIOCHEM ──
                using (SqlConnection conn = new SqlConnection(connectionString1))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DataRow row = dt.NewRow();
                            row["CT_Num"] = reader["CT_Num"]?.ToString();
                            row["CT_Intitule"] = reader["CT_Intitule"]?.ToString();
                            row["StockADate"] = Convert.ToDouble(reader["StockADate"]);
                            row["StockMinimum"] = Convert.ToDouble(reader["StockMinimum"]);
                            row["StockMaximum"] = Convert.ToDouble(reader["StockMaximum"]);
                            row["Source"] = "ARBIOCHEM";
                            dt.Rows.Add(row);
                        }
                    }
                }

                // ── ACTIVO ──
                using (SqlConnection conn = new SqlConnection(connectionString2))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DataRow row = dt.NewRow();
                            row["CT_Num"] = reader["CT_Num"]?.ToString();
                            row["CT_Intitule"] = reader["CT_Intitule"]?.ToString();
                            row["StockADate"] = Convert.ToDouble(reader["StockADate"]);
                            row["StockMinimum"] = Convert.ToDouble(reader["StockMinimum"]);
                            row["StockMaximum"] = Convert.ToDouble(reader["StockMaximum"]);
                            row["Source"] = "ACTIVO";
                            dt.Rows.Add(row);
                        }
                    }
                }
            }

            // ══════════════════════════════════════════════════════════════════════════════
            // 2. HELPER FORMATAGE NUMÉRIQUE
            // ══════════════════════════════════════════════════════════════════════════════

            string FormatNum(object val)
            {
                if (double.TryParse(val?.ToString(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double number))
                    return number.ToString("N2", new System.Globalization.CultureInfo("fr-FR"));
                return "0,00";
            }

            // ══════════════════════════════════════════════════════════════════════════════
            // 3. INTERFACE UTILISATEUR
            // ══════════════════════════════════════════════════════════════════════════════

            // ── Panel de recherche ──
            PanelControl searchPanel = new PanelControl();
            searchPanel.Dock = DockStyle.Top;
            searchPanel.Height = 60;
            searchPanel.BackColor = Color.FromArgb(240, 240, 240);
            searchPanel.Padding = new Padding(5);
            this.Controls.Add(searchPanel);

            LabelControl lblSearch = new LabelControl();
            lblSearch.Text = "Fournisseur :";
            lblSearch.Location = new Point(10, 12);
            lblSearch.Parent = searchPanel;

            TextEdit txtSearch = new TextEdit();
            txtSearch.Properties.NullValuePrompt = "🔍 Rechercher fournisseur";
            txtSearch.Properties.NullValuePromptShowForEmptyValue = true;
            txtSearch.Location = new Point(10, 30);
            txtSearch.Width = 180;
            txtSearch.Height = 24;
            txtSearch.Parent = searchPanel;

            LabelControl lblSearch1 = new LabelControl();
            lblSearch1.Text = "Article :";
            lblSearch1.Location = new Point(200, 12);
            lblSearch1.Parent = searchPanel;

            TextEdit txtSearch1 = new TextEdit();
            txtSearch1.Properties.NullValuePrompt = "🔍 Rechercher article";
            txtSearch1.Properties.NullValuePromptShowForEmptyValue = true;
            txtSearch1.Location = new Point(200, 30);
            txtSearch1.Width = 180;
            txtSearch1.Height = 24;
            txtSearch1.Parent = searchPanel;

            // ── Filtre BASE ──
            LabelControl lblBase = new LabelControl();
            lblBase.Text = "Base :";
            lblBase.Location = new Point(390, 12);
            lblBase.Parent = searchPanel;

            ComboBoxEdit cboBase = new ComboBoxEdit();
            cboBase.Properties.Items.AddRange(new[] { "Toutes", "ARBIOCHEM", "ACTIVO" });
            cboBase.EditValue = "Toutes";
            cboBase.Location = new Point(390, 30);
            cboBase.Width = 100;
            cboBase.Height = 24;
            cboBase.Parent = searchPanel;

            SimpleButton btnClear = new SimpleButton();
            btnClear.Text = "✕ Reset";
            btnClear.Location = new Point(500, 29);
            btnClear.Size = new Size(70, 24);
            btnClear.Parent = searchPanel;

            // ── Panel bas ──
            PanelControl bottomPanel = new PanelControl();
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.Height = 40;
            bottomPanel.BackColor = Color.FromArgb(45, 45, 48);

            LabelControl lblCount = new LabelControl();
            lblCount.Name = "lblCount";
            lblCount.Text = "0 article(s) sélectionné(s)";
            lblCount.ForeColor = Color.AliceBlue;
            lblCount.Font = new Font("Tahoma", 9f, FontStyle.Regular);
            lblCount.Location = new Point(10, 12);
            lblCount.Parent = bottomPanel;

            // ── Badge ARBIOCHEM ──
            LabelControl lblBadge1 = new LabelControl();
            lblBadge1.Text = "● ARBIOCHEM";
            lblBadge1.ForeColor = Color.FromArgb(100, 160, 230);
            lblBadge1.Font = new Font("Tahoma", 8f, FontStyle.Bold);
            lblBadge1.Location = new Point(300, 12);
            lblBadge1.Parent = bottomPanel;

            // ── Badge ACTIVO ──
            LabelControl lblBadge2 = new LabelControl();
            lblBadge2.Text = "● ACTIVO";
            lblBadge2.ForeColor = Color.FromArgb(230, 160, 60);
            lblBadge2.Font = new Font("Tahoma", 8f, FontStyle.Bold);
            lblBadge2.Location = new Point(380, 12);
            lblBadge2.Parent = bottomPanel;

            SimpleButton btnGenerer = new SimpleButton();
            btnGenerer.Text = "Générer PA";
            btnGenerer.Dock = DockStyle.Right;
            btnGenerer.Width = 160;
            btnGenerer.Appearance.BackColor = Color.FromArgb(39, 174, 96);
            btnGenerer.Appearance.ForeColor = Color.White;
            btnGenerer.Enabled = false;
            btnGenerer.Appearance.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            btnGenerer.Padding = new Padding(10, 5, 10, 5);
            btnGenerer.Parent = bottomPanel;

            // ══════════════════════════════════════════════════════════════════════════════
            // 4. GRILLE — LAYOUT VIEW
            // ══════════════════════════════════════════════════════════════════════════════

            LoadData();

            gdList.DataSource = null;
            layoutView = new LayoutView(gdList);
            layoutView.OptionsView.ViewMode = LayoutViewMode.MultiRow;
            gdList.MainView = layoutView;
            gdList.DataSource = dt;
            gdList.RefreshDataSource();

            layoutView.CardCaptionFormat = "{CT_Intitule}";
            layoutView.Appearance.CardCaption.Font = new Font("Tahoma", 8f, FontStyle.Bold);
            layoutView.Appearance.CardCaption.ForeColor = Color.White;
            layoutView.Appearance.CardCaption.BackColor = Color.FromArgb(255, 41, 98, 162);
            layoutView.Appearance.CardCaption.BorderColor = Color.FromArgb(255, 41, 98, 162);
            layoutView.CardMinSize = new Size(560, 300);

            // Masquer colonnes inutiles
            foreach (string colName in new[] { "CT_Intitule", "CT_Num", "StockADate",
                                    "StockMinimum", "StockMaximum", "Source" })
            {
                LayoutViewColumn c = layoutView.Columns[colName];
                if (c != null) c.Visible = false;
            }

            // Tri par CT_Intitule
            LayoutViewColumn colSort = layoutView.Columns["CT_Intitule"];
            if (colSort != null)
                layoutView.SortInfo.Add(new GridColumnSortInfo(colSort, ColumnSortOrder.Ascending));

            // ══════════════════════════════════════════════════════════════════════════════
            // 5. DESSIN PERSONNALISÉ DES CARTES
            // ══════════════════════════════════════════════════════════════════════════════

            layoutView.CustomDrawCardBackground += (sender, e) =>
            {
                e.DefaultDraw();
                cardBoundsDict[e.RowHandle] = e.Bounds;

                string ctNum = layoutView.GetRowCellValue(e.RowHandle, "CT_Num")?.ToString();
                string ctIntitule = layoutView.GetRowCellValue(e.RowHandle, "CT_Intitule")?.ToString() ?? "";
                string source = layoutView.GetRowCellValue(e.RowHandle, "Source")?.ToString() ?? "ARBIOCHEM";

                // ── Connexion selon la base ──
                string connStr = source == "ARBIOCHEM" ? connectionString1 : connectionString2;

                // ── Couleurs selon la base ──
                Color cardHeaderColor = source == "ARBIOCHEM"
                    ? Color.FromArgb(41, 98, 162)     // 🔵 Bleu  → ARBIOCHEM
                    : Color.FromArgb(162, 98, 41);    // 🟠 Orange → ACTIVO

                Color cardBorderColor = source == "ARBIOCHEM"
                    ? Color.FromArgb(100, 150, 220)
                    : Color.FromArgb(220, 150, 60);

                // ── Filtres texte ──
                string searchF = txtSearch?.Text?.Trim() ?? "";
                string searchA = txtSearch1?.Text?.Trim() ?? "";
                string baseFilter = cboBase?.EditValue?.ToString() ?? "Toutes";

                // ── Filtre base ──
                if (baseFilter == "ARBIOCHEM" && source != "ARBIOCHEM") return;
                if (baseFilter == "ACTIVO" && source != "ACTIVO") return;

                // ── Filtre fournisseur ──
                if (!string.IsNullOrEmpty(searchF) &&
                    ctIntitule.IndexOf(searchF, StringComparison.OrdinalIgnoreCase) < 0)
                    return;

                // ── Charger articles depuis la bonne base ──
                var articles = new List<(string Ref, string Design,
                                         double StockADate, double StockMin, double StockMax)>();

                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string sql = @"SELECT AR_Ref, AR_Design, StockADate, StockMinimum, StockMaximum 
                       FROM VW_SEUIL_MIN_STOCK_SUM 
                       WHERE CT_Num = @ct
                       ORDER BY AR_Design";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ct", ctNum);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                double ParseVal(string colName)
                                {
                                    string raw = reader[colName]?.ToString();
                                    if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                                        new System.Globalization.CultureInfo("fr-FR"), out double v1)) return v1;
                                    if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                                        System.Globalization.CultureInfo.InvariantCulture, out double v2)) return v2;
                                    return 0;
                                }

                                string arRef = reader["AR_Ref"]?.ToString();
                                articleConnectionMap[$"{ctNum}|{source}|{arRef}"] = connStr;

                                articles.Add((
                                    arRef,
                                    reader["AR_Design"]?.ToString(),
                                    ParseVal("StockADate"),
                                    ParseVal("StockMinimum"),
                                    ParseVal("StockMaximum")
                                ));
                            }
                        }
                    }
                }

                // ── Filtre article ──
                articles = articles
                    .Where(a => string.IsNullOrEmpty(searchA) ||
                           (a.Design ?? "").IndexOf(searchA, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                // ── Layout dimensions ──
                int headerH = 26;
                int colHeaderH = 22;
                int lineH = 22;
                int cbSize = 12;
                int maxVisible = 10;
                int headerTop = e.Bounds.Top + 4;
                int colHeaderTop = headerTop + headerH;
                int articleAreaTop = colHeaderTop + colHeaderH;
                int col1W = (e.Bounds.Width - cbSize - 24) * 2 / 3;
                int col2W = (e.Bounds.Width - cbSize - 24) / 3;

                using (Font fontHeader = new Font("Tahoma", 7.5f, FontStyle.Bold))
                using (Font fontCaption = new Font("Tahoma", 8.5f, FontStyle.Bold))
                using (Font fontArt = new Font("Tahoma", 7.5f))
                using (Font fontStock = new Font("Tahoma", 6.5f))
                using (Font fontBadge = new Font("Tahoma", 7f, FontStyle.Bold))
                using (SolidBrush fgWhite = new SolidBrush(Color.White))
                using (SolidBrush fgHeader = new SolidBrush(Color.Black))
                using (SolidBrush fgDesign = new SolidBrush(Color.FromArgb(40, 40, 40)))
                using (SolidBrush fgStock = new SolidBrush(Color.FromArgb(80, 80, 80)))
                using (SolidBrush fgMore = new SolidBrush(Color.Gray))
                {
                    var sfLeft = new StringFormat { LineAlignment = StringAlignment.Center };
                    var sfCenter = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    };
                    var sfRight = new StringFormat
                    {
                        Alignment = StringAlignment.Far,
                        LineAlignment = StringAlignment.Center
                    };

                    // ── En-tête carte colorée ──
                    using (SolidBrush headerBg = new SolidBrush(cardHeaderColor))
                        e.Cache.FillRectangle(headerBg,
                            new Rectangle(e.Bounds.Left + 1, headerTop, e.Bounds.Width - 2, headerH));

                    // ── Nom du fournisseur ──
                    e.Cache.DrawString(ctIntitule, fontCaption, fgWhite,
                        new Rectangle(e.Bounds.Left + 8, headerTop, e.Bounds.Width - 90, headerH),
                        sfLeft);

                    // ── Badge BASE ──
                    string badgeText = source == "ARBIOCHEM" ? "● ARBIOCHEM" : "● ACTIVO";
                    Color badgeBgColor = source == "ARBIOCHEM"
                        ? Color.FromArgb(200, 20, 70, 140)
                        : Color.FromArgb(200, 140, 70, 20);

                    Rectangle badgeRect = new Rectangle(e.Bounds.Right - 72, headerTop + 4, 68, 18);

                    using (SolidBrush badgeBg = new SolidBrush(badgeBgColor))
                        e.Cache.FillRectangle(badgeBg, badgeRect);

                    e.Cache.DrawString(badgeText, fontBadge, fgWhite, badgeRect, sfCenter);

                    // ── En-tête colonnes ──
                    using (SolidBrush colHeaderBg = new SolidBrush(Color.FromArgb(220, 220, 220)))
                        e.Cache.FillRectangle(colHeaderBg,
                            new Rectangle(e.Bounds.Left + 1, colHeaderTop, e.Bounds.Width - 2, colHeaderH));

                    e.Cache.DrawString("Désignation article", fontHeader, fgHeader,
                        new Rectangle(e.Bounds.Left + cbSize + 14, colHeaderTop, col1W, colHeaderH), sfLeft);

                    e.Cache.DrawString("A date | Min | Max", fontHeader, fgHeader,
                        new Rectangle(e.Bounds.Left + cbSize + 14 + col1W, colHeaderTop, col2W, colHeaderH), sfRight);

                    // ── Séparateur colonne en-tête ──
                    using (Pen sepHeader = new Pen(Color.FromArgb(180, 180, 180)))
                    {
                        int xSepH = e.Bounds.Left + cbSize + 14 + col1W;
                        e.Cache.DrawLine(sepHeader,
                            new Point(xSepH, colHeaderTop),
                            new Point(xSepH, colHeaderTop + colHeaderH));
                    }

                    // ── Lignes articles ──
                    int drawn = 0;

                    foreach (var (arRef, arDesign, stockADate, stockMin, stockMax) in articles)
                    {
                        if (drawn >= maxVisible) break;

                        int y = articleAreaTop + drawn * lineH;

                        // Couleur ligne
                        Color baseColor = (drawn % 2 == 0)
                            ? Color.FromArgb(248, 248, 248)
                            : Color.White;

                        Color lineColor = (stockADate < stockMin)
                            ? Color.FromArgb(160, 255, 80, 80)   // 🔴 stock insuffisant
                            : baseColor;

                        Rectangle lineRect = new Rectangle(
                            e.Bounds.Left + 1, y, e.Bounds.Width - 2, lineH);

                        using (SolidBrush bg = new SolidBrush(lineColor))
                            e.Cache.FillRectangle(bg, lineRect);

                        // ── Clé ligne = "rowHandle|AR_Ref|Source" ──
                        string lineKey = $"{e.RowHandle}|{arRef}|{source}";
                        articleLineBounds[lineKey] = lineRect;

                        // ── Clé article = "CT_Num|Source|AR_Ref" ──
                        string articleKey = $"{ctNum}|{source}|{arRef}";
                        if (!checkStates.ContainsKey(articleKey))
                            checkStates[articleKey] = false;

                        bool isChecked = checkStates[articleKey];

                        // ── Checkbox ──
                        Rectangle cbRect = new Rectangle(
                            e.Bounds.Left + 6,
                            y + (lineH - cbSize) / 2,
                            cbSize, cbSize);

                        using (SolidBrush cbBg = new SolidBrush(isChecked ? cardHeaderColor : Color.White))
                            e.Cache.FillRectangle(cbBg, cbRect);

                        using (Pen pen = new Pen(isChecked ? cardHeaderColor : Color.Gray))
                            e.Cache.DrawRectangle(pen, cbRect);

                        if (isChecked)
                        {
                            using (Pen checkPen = new Pen(Color.White, 2f))
                            {
                                e.Cache.DrawLine(checkPen,
                                    new Point(cbRect.Left + 2, cbRect.Top + cbSize / 2),
                                    new Point(cbRect.Left + cbSize / 2 - 1, cbRect.Bottom - 2));
                                e.Cache.DrawLine(checkPen,
                                    new Point(cbRect.Left + cbSize / 2 - 1, cbRect.Bottom - 2),
                                    new Point(cbRect.Right - 1, cbRect.Top + 2));
                            }
                        }

                        // ── Texte désignation ──
                        e.Cache.DrawString(arDesign ?? "", fontArt, fgDesign,
                            new Rectangle(e.Bounds.Left + cbSize + 14, y, col1W, lineH), sfLeft);

                        // ── Stock info ──
                        string stockInfo =
                            $"{FormatNum(stockADate),9} | {FormatNum(stockMin),9} | {FormatNum(stockMax),9}";

                        // Couleur stock rouge si insuffisant
                        using (SolidBrush fgStockColor = new SolidBrush(
                            stockADate < stockMin ? Color.DarkRed : Color.FromArgb(80, 80, 80)))
                        {
                            e.Cache.DrawString(stockInfo, fontStock, fgStockColor,
                                new Rectangle(e.Bounds.Left + cbSize + 14 + col1W, y, col2W, lineH), sfRight);
                        }

                        // ── Séparateur colonne ──
                        using (Pen sep = new Pen(Color.FromArgb(200, 200, 200)))
                        {
                            int xSep = e.Bounds.Left + cbSize + 14 + col1W;
                            e.Cache.DrawLine(sep, new Point(xSep, y), new Point(xSep, y + lineH));
                        }

                        // ── Séparateur ligne ──
                        using (Pen lineSep = new Pen(Color.FromArgb(230, 230, 230)))
                        {
                            e.Cache.DrawLine(lineSep,
                                new Point(e.Bounds.Left + 1, y + lineH - 1),
                                new Point(e.Bounds.Right - 1, y + lineH - 1));
                        }

                        drawn++;
                    }

                    // ── "+N articles..." ──
                    if (articles.Count > maxVisible)
                    {
                        int remain = articles.Count - maxVisible;
                        int yMore = articleAreaTop + maxVisible * lineH + 2;

                        e.Cache.DrawString(
                            $"  +{remain} article(s) non affichés...",
                            fontStock, fgMore,
                            new Rectangle(e.Bounds.Left + 10, yMore, e.Bounds.Width - 20, 16),
                            sfLeft);
                    }

                    // ── Bordure carte colorée ──
                    using (Pen borderPen = new Pen(cardBorderColor, 1.5f))
                        e.Cache.DrawRectangle(borderPen,
                            new Rectangle(e.Bounds.Left + 1, e.Bounds.Top + 1,
                                          e.Bounds.Width - 3, e.Bounds.Height - 3));
                }
            };

            // ══════════════════════════════════════════════════════════════════════════════
            // 6. GESTION DU CLIC SOURIS
            // ══════════════════════════════════════════════════════════════════════════════

            gdList.MouseClick += (sender, e) =>
            {
                Point pt = e.Location;

                foreach (var kvp in articleLineBounds)
                {
                    if (!kvp.Value.Contains(pt)) continue;

                    // clé = "rowHandle|AR_Ref|Source"
                    string[] parts = kvp.Key.Split('|');
                    if (parts.Length < 3) continue;

                    int rowHandle = int.Parse(parts[0]);
                    string arRef = parts[1];
                    string source = parts[2];

                    string ctNum2 = layoutView.GetRowCellValue(rowHandle, "CT_Num")?.ToString();
                    string articleKey = $"{ctNum2}|{source}|{arRef}";

                    if (!checkStates.ContainsKey(articleKey))
                        checkStates[articleKey] = false;

                    checkStates[articleKey] = !checkStates[articleKey];

                    UpdateSelectionCount();
                    layoutView.RefreshData();
                    return;
                }
            };

            // ══════════════════════════════════════════════════════════════════════════════
            // 7. FILTRES — RECHERCHE
            // ══════════════════════════════════════════════════════════════════════════════

            txtSearch.TextChanged += (s, ev) =>
            {
                cardBoundsDict.Clear();
                articleLineBounds.Clear();
                layoutView.RefreshData();
            };

            txtSearch1.TextChanged += (s, ev) =>
            {
                cardBoundsDict.Clear();
                articleLineBounds.Clear();
                layoutView.RefreshData();
            };

            cboBase.EditValueChanged += (s, ev) =>
            {
                string baseFilter = cboBase.EditValue?.ToString() ?? "Toutes";

                try
                {
                    if (baseFilter == "Toutes")
                        dt.DefaultView.RowFilter = string.Empty;
                    else if (baseFilter == "ARBIOCHEM")
                        dt.DefaultView.RowFilter = "Source = 'ARBIOCHEM'";
                    else if (baseFilter == "ACTIVO")
                        dt.DefaultView.RowFilter = "Source = 'ACTIVO'";

                    gdList.DataSource = null;
                    gdList.DataSource = dt.DefaultView;
                    gdList.MainView = layoutView;
                    gdList.RefreshDataSource();
                    cardBoundsDict.Clear();
                    articleLineBounds.Clear();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur filtre base : " + ex.Message);
                }
            };

            txtSearch.EditValueChanged += (sender, e) =>
            {
                string search = txtSearch.Text.Trim();
                try
                {
                    string baseFilter = cboBase.EditValue?.ToString() ?? "Toutes";
                    string baseCondition = baseFilter == "ARBIOCHEM" ? "Source = 'ARBIOCHEM'" :
                                           baseFilter == "ACTIVO" ? "Source = 'ACTIVO'" : "";

                    string searchCondition = string.IsNullOrEmpty(search) ? ""
                        : $"CT_Intitule LIKE '%{search.Replace("'", "''")}%'";

                    string filter = string.Join(" AND ",
                        new[] { baseCondition, searchCondition }.Where(c => !string.IsNullOrEmpty(c)));

                    dt.DefaultView.RowFilter = filter;
                    gdList.DataSource = null;
                    gdList.DataSource = dt.DefaultView;
                    gdList.MainView = layoutView;
                    gdList.RefreshDataSource();
                    cardBoundsDict.Clear();
                    articleLineBounds.Clear();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur filtre : " + ex.Message);
                }
            };

            // ── Bouton reset ──
            btnClear.Click += (sender, e) =>
            {
                txtSearch.EditValue = null;
                txtSearch1.EditValue = null;
                cboBase.EditValue = "";
                dt.DefaultView.RowFilter = string.Empty;
                gdList.DataSource = null;
                gdList.DataSource = dt;
                gdList.MainView = layoutView;
                gdList.RefreshDataSource();
                cardBoundsDict.Clear();
                articleLineBounds.Clear();
            };

            // ══════════════════════════════════════════════════════════════════════════════
            // 8. COMPTEUR DE SÉLECTION
            // ══════════════════════════════════════════════════════════════════════════════

            void UpdateSelectionCount()
            {
                int count = checkStates.Count(kv => kv.Value);
                int countBase1 = checkStates.Count(kv => kv.Value && kv.Key.Split('|')[1] == "ARBIOCHEM");
                int countBase2 = checkStates.Count(kv => kv.Value && kv.Key.Split('|')[1] == "ACTIVO");

                lblCount.Text = count == 0
                    ? "0 article(s) sélectionné(s)"
                    : $"{count} article(s)  |  ARBIOCHEM: {countBase1}  |  ACTIVO: {countBase2}";

                btnGenerer.Enabled = count > 0;
            }

            // ══════════════════════════════════════════════════════════════════════════════
            // 9. GÉNÉRATION PA — 2 BASES SÉPARÉES
            // ══════════════════════════════════════════════════════════════════════════════

            btnGenerer.Click += (sender, e) =>
            {
                // ── Extraire articles cochés avec leur source ──
                var selectedKeys = checkStates
                    .Where(kv => kv.Value)
                    .Select(kv =>
                    {
                        var p = kv.Key.Split('|');
                        return (CT_Num: p[0], Source: p[1], AR_Ref: p[2]);
                    })
                    .ToList();

                if (selectedKeys.Count == 0)
                {
                    MessageBox.Show("Aucun article sélectionné.", "Attention",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // ── Grouper par CT_Num + Source (une carte = un groupe) ──
                var groupedByFournisseur = selectedKeys
                    .GroupBy(x => new { x.CT_Num, x.Source })
                    .Select(g =>
                    {
                        string ctNum = g.Key.CT_Num;
                        string source = g.Key.Source;
                        string connStr = source == "ARBIOCHEM" ? connectionString1 : connectionString2;
                        var arRefs = g.Select(x => x.AR_Ref).ToList();

                        // Récupérer le nom du fournisseur
                        DataRow fournisseurRow = dt.Rows.Cast<DataRow>()
                            .FirstOrDefault(r =>
                                r["CT_Num"]?.ToString() == ctNum &&
                                r["Source"]?.ToString() == source);

                        string fournisseur = fournisseurRow?["CT_Intitule"]?.ToString() ?? ctNum;

                        // Charger articles cochés depuis la bonne base
                        DataTable dtArticles = new DataTable();
                        using (var conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            string sql = @"SELECT AR_Ref, AR_Design, StockADate, StockMinimum, StockMaximum
                               FROM VW_SEUIL_MIN_STOCK_SUM
                               WHERE CT_Num = @ctNum";
                            var ada = new SqlDataAdapter(sql, conn);
                            ada.SelectCommand.Parameters.AddWithValue("@ctNum", ctNum);
                            ada.Fill(dtArticles);
                        }

                        // Filtrer uniquement les articles cochés
                        DataTable dtFiltered = dtArticles.Clone();
                        foreach (DataRow row in dtArticles.Rows)
                        {
                            string arRef = row["AR_Ref"]?.ToString();
                            string articleKey = $"{ctNum}|{source}|{arRef}";
                            if (checkStates.ContainsKey(articleKey) && checkStates[articleKey])
                                dtFiltered.ImportRow(row);
                        }

                        return (CT_Num: ctNum, Source: source, ConnStr: connStr,
                                Fournisseur: fournisseur, Articles: dtFiltered);
                    })
                    .Where(g => g.Articles.Rows.Count > 0)
                    .ToList();

                if (groupedByFournisseur.Count == 0)
                {
                    MessageBox.Show("Aucun article sélectionné.", "Attention",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // ── Récapitulatif ──
                var recap = new StringBuilder();
                recap.AppendLine($"{groupedByFournisseur.Count} PA(s) vont être générés :\n");

                foreach (var (ctNum, source, connStr, fournisseur, articles) in groupedByFournisseur)
                {
                    string baseLabel = source == "ARBIOCHEM" ? "🔵 ARBIOCHEM" : "🟠 ACTIVO";
                    recap.AppendLine($"📦 {fournisseur}  [{baseLabel}]  ({articles.Rows.Count} article(s))");

                    foreach (DataRow r in articles.Rows)
                        recap.AppendLine($"   • {r["AR_Design"]}");

                    recap.AppendLine();
                }

                DialogResult confirm = MessageBox.Show(recap.ToString(), "Confirmation",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                if (confirm != DialogResult.OK) return;

                // ── Génération ──
                int paGeneresCount = 0;
                var erreurs = new List<string>();
                int depot = 0, cono = 0;

                foreach (var (ctNum, source, connStr, fournisseur, articles) in groupedByFournisseur)
                {
                    doPiece = GetNextInvoiceNumber(_prefix);

                    try
                    {
                        // EN-TÊTE
                        using (var connection = new SqlConnection(connectionStrings))
                        {
                            connection.Open();
                            var triggers = new[] { "TG_CBINS_F_DOCENTETE", "TG_INS_CPTAF_DOCENTETE", "TG_INS_F_DOCENTETE" };

                            try
                            {
                                foreach (var t in triggers)
                                    new SqlCommand($"DISABLE TRIGGER {t} ON F_DOCENTETE", connection).ExecuteNonQuery();

                                dodo_piece = doPiece;
                                InsertFDOCENTETE(ctNum, fournisseur, connectionStrings);
                            }
                            finally
                            {
                                foreach (var t in triggers)
                                    new SqlCommand($"ENABLE TRIGGER {t} ON F_DOCENTETE", connection).ExecuteNonQuery();
                            }
                        }

                        decimal cours = 0, montant_total = 0;

                        // LIGNES
                        foreach (DataRow article in articles.Rows)
                        {
                            string arRef = article["AR_Ref"]?.ToString();
                            string arDesign = article["AR_Design"]?.ToString();

                            using (var frmQte = new frm_saisieQuantite(arRef, arDesign, dodo_piece, fournisseur))
                            {
                                if (frmQte.ShowDialog() != DialogResult.OK) continue;

                                int qte = int.Parse(frmQte.txtQte.Text);
                                decimal pu = decimal.Parse(frmQte.txtKg.Text);
                                cours = decimal.Parse(frmQte.txtcoursdevise.Text);
                                devise = frmQte.cmbDevise.Text.ToString();
                                decimal montant = qte * pu * cours;
                                montant_total += montant;

                                depot = frmQte.deno;
                                cono = frmQte.cono;
                                dts = frmQte.dt;
                                dtc = frmQte.dc;

                                decimal poids = Convert.ToDecimal(frmQte.txtPoids.Text);

                                using (var connection = new SqlConnection(connectionStrings))
                                {
                                    connection.Open();
                                    var triggers = new[] { "TG_CBINS_F_DOCLIGNE", "TG_INS_CPTAF_DOCLIGNE", "TG_INS_F_DOCLIGNE" };

                                    try
                                    {
                                        foreach (var t in triggers)
                                            new SqlCommand($"DISABLE TRIGGER {t} ON F_DOCLIGNE", connection).ExecuteNonQuery();

                                        string sqlUpdate = @"UPDATE F_DOCENTETE 
                                        SET DO_Taxe1=0, DO_Expedit=0,DO_Devise=@devise,
                                            DE_No=@deno, CO_No=@cono, DO_Imprim=0, DO_Statut=0,
                                            DO_Cours=@doCours, DO_TotalHTNet=@montant,
                                            DO_TotalHT=@montant, DO_TotalTTC=@montant,
                                            DO_DateExpedition=@dtc
                                        WHERE DO_Piece = @doPiece";

                                        using (var cmd = new SqlCommand(sqlUpdate, connection))
                                        {
                                            cmd.Parameters.Add("@doCours", SqlDbType.Decimal).Value = cours;
                                            cmd.Parameters.Add("@doPiece", SqlDbType.Char, 13).Value = dodo_piece.PadRight(13);
                                            cmd.Parameters.Add("@montant", SqlDbType.Decimal).Value = montant_total;
                                            cmd.Parameters.Add("@deno", SqlDbType.Int).Value = depot;
                                            cmd.Parameters.Add("@cono", SqlDbType.Int).Value = cono;
                                            cmd.Parameters.Add("@dtc", SqlDbType.DateTime).Value = dtc;
                                            cmd.Parameters.Add("@devise", SqlDbType.Int).Value = recuperer_devise(devise);
                                            cmd.ExecuteNonQuery();
                                        }

                                        InsertFDOCLIGNE(dodo_piece, arRef, arDesign, qte, pu,
                                                        montant, ctNum, depot, poids,connStr);
                                    }
                                    finally
                                    {
                                        foreach (var t in triggers)
                                            new SqlCommand($"ENABLE TRIGGER {t} ON F_DOCLIGNE", connection).ExecuteNonQuery();
                                    }
                                }
                            }
                        }

                        paGeneresCount++;
                    }
                    catch (Exception ex)
                    {
                        erreurs.Add($"• {fournisseur} [{source}] : {ex.Message}");
                    }
                }

                // ── Résultat final ──
                if (erreurs.Count > 0)
                {
                    MessageBox.Show(
                        $"{paGeneresCount} PA(s) générés.\n\nErreurs :\n{string.Join("\n", erreurs)}",
                        "Résultat partiel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(
                        $"✅ {paGeneresCount} PA(s) générés avec succès.",
                        "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                // ── Reset sélection ──
                checkStates.Clear();
                articleConnectionMap.Clear();
                UpdateSelectionCount();
                layoutView.RefreshData();
            };

            // ══════════════════════════════════════════════════════════════════════════════
            // 10. AJOUT CONTRÔLES + REFRESH
            // ══════════════════════════════════════════════════════════════════════════════

            this.Controls.Add(bottomPanel);
            this.Controls.Add(searchPanel);
            searchPanel.BringToFront();

            layoutView.RefreshData();
        }

        private int recuperer_devise(String cond)
        {
            AppDbContext context = new AppDbContext();
            short d_val = 0;

            var test = context.P_DEVISE.Where(p => p.D_Intitule.Contains(cond)).FirstOrDefault();

            if(test != null)
            {
                d_val = (short)test.cbIndice;
            }
            return d_val;
        }
    }
}