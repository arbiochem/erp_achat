using arbioApp.DTO;
using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Services;
using DevExpress.Data;
using DevExpress.DataProcessing;
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

        public static string recuperer_ctnum(string fournisseur)
        {
            string query = "SELECT CT_Num FROM F_COMPTET WHERE CT_Intitule = @intitule";

            // Récupérer le dernier numéro
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.Add("@intitule", SqlDbType.VarChar).Value = fournisseur;

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
        private string InsertFDOCENTETE(string ct_Num, string fournisseur)
        {
            using (SqlConnection conn = new SqlConnection(connectionStrings))
            {
                conn.Open();
                ct_Num = recuperer_ctnum(fournisseur);

                string sql = @"
                    INSERT INTO F_DOCENTETE
                        (DO_Domaine,DO_Type, DO_Piece, DO_Date,DO_Tiers)
                    VALUES
                        (1,10, @doPiece, @doDate,@dotiers)";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@doPiece", SqlDbType.Char, 13).Value = doPiece.PadRight(13);
                    cmd.Parameters.Add("@dotiers", SqlDbType.Char, 13).Value = ct_Num.PadRight(13);

                    // DO_Date est DateTime dans Sage 100
                    cmd.Parameters.Add("@doDate", SqlDbType.DateTime).Value = DateTime.Today;

                    cmd.ExecuteNonQuery();
                }

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
                                      double qteACommander = 0,decimal pu=0,decimal montant=0,string ct_num="",int depot=0,decimal poids=0)
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
                string sql = @"
                INSERT INTO F_DOCLIGNE
                    (DO_Domaine, DO_Type, DO_Piece, DL_Ligne, AR_Ref, DL_Design, PF_Num, DL_No,DL_Qte,DL_PrixUnitaire,DL_CMUP,DL_PUTTC,DL_MontantHT,DL_MontantTTC,CT_Num,DE_No,cbDE_No,DO_Date,DL_DateBC,DL_DateBL,DL_PoidsNet,cbCreationUser,retenu)
                VALUES
                    (1, 10, @doPiece, @dlLigne, @arRef, @design, @pfNum, @dlNo, @dlqte,@dlpu,@dlpu,@dlpu,@montant,@montant,@ctnum,@deno,@deno,@date,@date,@date,@poids,@user,0)";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    int dlNo = GetNextIDLNo(conn); // ← unique à chaque ligne

                    cmd.Parameters.Add("@doPiece", SqlDbType.Char, 13).Value = doPiece.PadRight(13);
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

        public void ChargerDonneesDepuisBDD()
        {
            // ── Dictionnaires d'état ──
            // articleLineBounds : clé = "rowHandle|AR_Ref", valeur = Rectangle de la ligne
            // checkStates       : clé = "CT_Num|AR_Ref",   valeur = bool coché
            // cardBoundsDict    : clé = rowHandle,          valeur = Rectangle de la carte

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

            SimpleButton btnGenerer = new SimpleButton();
            btnGenerer.Text = "Générer PA";
            btnGenerer.Dock = DockStyle.Right;
            btnGenerer.Width = 160;
            btnGenerer.Appearance.BackColor = Color.FromArgb(39, 174, 96);
            btnGenerer.Appearance.ForeColor = Color.White;
            btnGenerer.Enabled = false;
            btnGenerer.Appearance.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            btnGenerer.Padding = new Padding(10, 5, 10, 5);
            btnGenerer.Margin = new Padding(5, 5, 10, 5);
            btnGenerer.Parent = bottomPanel;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = @"
                SELECT 
                    CT_Num,
                    CT_Intitule,
                    SUM(StockADate)   AS StockADate,
                    MIN(StockMinimum) AS StockMinimum,
                    MAX(StockMaximum) AS StockMaximum
                FROM VW_SEUIL_MIN_STOCK_SUM
                GROUP BY CT_Num, CT_Intitule";
                SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                dt = new DataTable();
                adapter.Fill(dt);
            }

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

            LayoutViewColumn col = layoutView.Columns["CT_Intitule"];
            if (col != null)
            {
                col.Visible = false;
                layoutView.SortInfo.Add(new GridColumnSortInfo(col, ColumnSortOrder.Ascending));
            }

            foreach (string colName in new[] { "Dateop", "CT_Num", "StockADate", "StockMinimum", "StockMaximum" })
            {
                LayoutViewColumn c = layoutView.Columns[colName];
                if (c != null) c.Visible = false;
            }

            layoutView.CardMinSize = new Size(550, 280);

            // ── Helper formatage numérique ──
            string FormatNum(object val)
            {
                if (double.TryParse(val?.ToString(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double number))
                    return number.ToString("N2", new System.Globalization.CultureInfo("fr-FR"));
                return "0,00";
            }

            // ── Panels de recherche ──
            Panel spacer = new Panel();
            spacer.Dock = DockStyle.Top;
            spacer.Height = 10;
            this.Controls.Add(spacer);

            PanelControl searchPanel = new PanelControl();
            searchPanel.Dock = DockStyle.Top;
            searchPanel.Height = 60;
            searchPanel.BackColor = Color.FromArgb(240, 240, 240);
            searchPanel.Padding = new Padding(5);
            this.Controls.Add(searchPanel);

            LabelControl lblSearch = new LabelControl();
            lblSearch.Text = "Recherche :";
            lblSearch.Location = new Point(10, 12);
            lblSearch.Parent = searchPanel;

            TextEdit txtSearch = new TextEdit();
            txtSearch.Properties.NullValuePrompt = "🔍 Fournisseur";
            txtSearch.Properties.NullValuePromptShowForEmptyValue = true;
            txtSearch.Location = new Point(10, 30);
            txtSearch.Width = 140;
            txtSearch.Height = 24;
            txtSearch.Parent = searchPanel;

            TextEdit txtSearch1 = new TextEdit();
            txtSearch1.Properties.NullValuePrompt = "🔍 Article";
            txtSearch1.Properties.NullValuePromptShowForEmptyValue = true;
            txtSearch1.Location = new Point(160, 30);
            txtSearch1.Width = 140;
            txtSearch1.Height = 24;
            txtSearch1.Parent = searchPanel;

            SimpleButton btnClear = new SimpleButton();
            btnClear.Text = "✕";
            btnClear.Location = new Point(305, 29);
            btnClear.Size = new Size(30, 24);
            btnClear.Parent = searchPanel;

            txtSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtSearch1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // ── Rafraîchissement sur saisie dans les deux champs ──
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

            // ── Dessin personnalisé des cartes ──
            layoutView.CustomDrawCardBackground += (sender, e) =>
            {
                e.DefaultDraw();
                cardBoundsDict[e.RowHandle] = e.Bounds;

                string ctNum = layoutView.GetRowCellValue(e.RowHandle, "CT_Num")?.ToString();

                // ✅ CT_Intitule récupéré directement depuis l'en-tête de la carte (DataTable)
                string ctIntitule = layoutView.GetRowCellValue(e.RowHandle, "CT_Intitule")?.ToString() ?? "";

                // ── Charger les articles depuis la BDD ──
                var articles = new List<(string client, string Ref, string Design, double StockADate, double StockMin, double StockMax)>();

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string sql = @"SELECT CT_Intitule, AR_Ref, AR_Design, StockADate, StockMinimum, StockMaximum 
                   FROM VW_SEUIL_MIN_STOCK_SUM WHERE CT_Num = @ct";

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

                                articles.Add((
                                    reader["CT_Intitule"]?.ToString(),
                                    reader["AR_Ref"]?.ToString(),
                                    reader["AR_Design"]?.ToString(),
                                    ParseVal("StockADate"),
                                    ParseVal("StockMinimum"),
                                    ParseVal("StockMaximum")
                                ));
                            }
                        }
                    }
                }

                // ── Filtre texte ──
                string searchF = txtSearch?.Text?.Trim() ?? "";
                string searchA = txtSearch1?.Text?.Trim() ?? "";

                // ✅ Le filtre fournisseur porte sur ctIntitule (en-tête de carte) — pas sur a.client
                if (!string.IsNullOrEmpty(searchF) &&
                    ctIntitule.IndexOf(searchF, StringComparison.OrdinalIgnoreCase) < 0)
                    return;

                // Filtre article sur la désignation uniquement
                articles = articles
                    .Where(a =>
                        string.IsNullOrEmpty(searchA) ||
                        (a.Design ?? "").IndexOf(searchA, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                // ── Layout ──
                int headerTop = e.Bounds.Top + 28;
                int headerH = 22;
                int articleAreaTop = headerTop + headerH;
                int articleAreaBottom = e.Bounds.Bottom - 4;

                int lineH = 22;
                int cbSize = 12;
                int maxVisible = 10;

                int col1W = (e.Bounds.Width - cbSize - 20) * 2 / 3;
                int col2W = (e.Bounds.Width - cbSize - 20) / 3;

                using (Font fontHeader = new Font("Tahoma", 7.5f, FontStyle.Bold))
                using (Font fontArt = new Font("Tahoma", 7.5f))
                using (Font fontStock = new Font("Tahoma", 6.5f))
                using (SolidBrush fgHeader = new SolidBrush(Color.Black))
                using (SolidBrush fgDesign = new SolidBrush(Color.FromArgb(40, 40, 40)))
                using (SolidBrush fgStock = new SolidBrush(Color.FromArgb(80, 80, 80)))
                using (SolidBrush fgMore = new SolidBrush(Color.Gray))
                {
                    var sfLeft = new StringFormat { LineAlignment = StringAlignment.Center };
                    var sfRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };

                    // ── En-tête colonnes ──
                    using (SolidBrush headerBg = new SolidBrush(Color.FromArgb(230, 230, 230)))
                        e.Cache.FillRectangle(headerBg,
                            new Rectangle(e.Bounds.Left + 2, headerTop, e.Bounds.Width - 4, headerH));

                    e.Cache.DrawString("Désignation", fontHeader, fgHeader,
                        new Rectangle(e.Bounds.Left + cbSize + 12, headerTop, col1W, headerH), sfLeft);

                    e.Cache.DrawString("A date | Min | Max", fontHeader, fgHeader,
                        new Rectangle(e.Bounds.Left + cbSize + 12 + col1W, headerTop, col2W, headerH), sfRight);

                    int drawn = 0;

                    foreach (var (client, arRef, arDesign, stockADate, stockMin, stockMax) in articles)
                    {
                        if (drawn >= maxVisible) break;

                        int y = articleAreaTop + drawn * lineH;

                        // ── Couleur de ligne ──
                        Color baseColor = (drawn % 2 == 0)
                            ? Color.FromArgb(245, 245, 245)
                            : Color.White;

                        Color lineColor = (stockADate < stockMin)
                            ? Color.FromArgb(180, 255, 80, 80)
                            : baseColor;

                        Rectangle lineRect = new Rectangle(e.Bounds.Left + 2, y, e.Bounds.Width - 4, lineH);

                        using (SolidBrush bg = new SolidBrush(lineColor))
                            e.Cache.FillRectangle(bg, lineRect);

                        // clé = "rowHandle|AR_Ref"  → utilisée dans MouseClick
                        string lineKey = $"{e.RowHandle}|{arRef}";
                        articleLineBounds[lineKey] = lineRect;

                        // ── Checkbox ──
                        string articleKey = $"{ctNum}|{arRef}";
                        if (!checkStates.ContainsKey(articleKey))
                            checkStates[articleKey] = false;

                        bool isChecked = checkStates[articleKey];

                        Rectangle cbRect = new Rectangle(
                            e.Bounds.Left + 6,
                            y + (lineH - cbSize) / 2,
                            cbSize, cbSize);

                        using (SolidBrush cbBg = new SolidBrush(isChecked ? Color.Green : Color.White))
                            e.Cache.FillRectangle(cbBg, cbRect);

                        using (Pen pen = new Pen(Color.Gray))
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

                        // ── Texte article ──
                        e.Cache.DrawString(arDesign ?? "", fontArt, fgDesign,
                            new Rectangle(e.Bounds.Left + cbSize + 12, y, col1W, lineH), sfLeft);

                        string stockInfo = $"{FormatNum(stockADate),9} | {FormatNum(stockMin),9} | {FormatNum(stockMax),9}";

                        e.Cache.DrawString(stockInfo, fontStock, fgStock,
                            new Rectangle(e.Bounds.Left + cbSize + 12 + col1W, y, col2W, lineH), sfRight);

                        // ── Séparateur colonne ──
                        using (Pen sep = new Pen(Color.FromArgb(200, 200, 200)))
                        {
                            int xSep = e.Bounds.Left + cbSize + 12 + col1W;
                            e.Cache.DrawLine(sep, new Point(xSep, y), new Point(xSep, y + lineH));
                        }

                        drawn++;
                    }

                    // ── "+N articles..." si tronqué ──
                    if (articles.Count > maxVisible)
                    {
                        int remain = articles.Count - maxVisible;
                        e.Cache.DrawString($"+{remain} article(s)...", fontStock, fgMore,
                            new Rectangle(e.Bounds.Left + 10, articleAreaBottom - 16,
                                          e.Bounds.Width - 20, 16), sfLeft);
                    }
                }
            };

            // ── Gestion du clic souris ──
            gdList.MouseClick += (sender, e) =>
            {
                Point pt = e.Location;

                // Parcourir articleLineBounds pour détecter le clic
                foreach (var kvp in articleLineBounds)
                {
                    if (kvp.Value.Contains(pt))
                    {
                        // clé = "rowHandle|AR_Ref"
                        string[] parts = kvp.Key.Split('|');
                        if (parts.Length < 2) continue;

                        int rowHandle = int.Parse(parts[0]);
                        string arRef = parts[1];
                        string ctNum2 = layoutView.GetRowCellValue(rowHandle, "CT_Num")?.ToString();

                        string articleKey = $"{ctNum2}|{arRef}";

                        if (!checkStates.ContainsKey(articleKey))
                            checkStates[articleKey] = false;

                        checkStates[articleKey] = !checkStates[articleKey];

                        UpdateSelectionCount();
                        layoutView.RefreshData();
                        return;
                    }
                }
            };

            this.Controls.Add(searchPanel);
            searchPanel.BringToFront();
            this.Controls.Add(bottomPanel);

            // ── Mise à jour du compteur de sélection ──
            void UpdateSelectionCount()
            {
                int count = checkStates.Count(kv => kv.Value);
                lblCount.Text = $"{count} article(s) sélectionné(s)";
                btnGenerer.Enabled = count > 0;
            }

            // ── Filtre sur fournisseur (RowFilter sur DataTable) ──
            txtSearch.EditValueChanged += (sender, e) =>
            {
                string search = txtSearch.Text.Trim();
                try
                {
                    if (string.IsNullOrEmpty(search))
                        dt.DefaultView.RowFilter = string.Empty;
                    else
                    {
                        string safe = search.Replace("'", "''");
                        dt.DefaultView.RowFilter = $"CT_Intitule LIKE '%{safe}%'";
                    }

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
                dt.DefaultView.RowFilter = string.Empty;
                gdList.DataSource = null;
                gdList.DataSource = dt;
                gdList.MainView = layoutView;
                gdList.RefreshDataSource();
                cardBoundsDict.Clear();
                articleLineBounds.Clear();
            };
        

            btnGenerer.Click += (sender, e) =>
            {

                // 1. Fournisseurs cochés
                // Récupère les CT_Num distincts des articles cochés
                var selectedCtNums = checkStates
                    .Where(kv => kv.Value)
                    .Select(kv => kv.Key.Split('|')[0]) // CT_Num
                    .Distinct()
                    .ToList();

                if (selectedCtNums.Count == 0)
                {
                    MessageBox.Show("Aucun article sélectionné.", "Attention",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

               
                // 1. Afficher les fournisseurs concernés
                var selectedFournisseurs = dt.Rows
                    .Cast<DataRow>()
                    .Where(r => selectedCtNums.Contains(r["CT_Num"]?.ToString()))
                    .Select(r => r["CT_Intitule"]?.ToString())
                    .Distinct()
                    .ToList();

                // 2. Charger + filtrer articles cochés
                var groupedByFournisseur = new List<(string CT_Num, string Fournisseur, DataTable Articles)>();

                foreach (string ctNum in selectedCtNums)
                {
                    DataRow fournisseurRow = dt.Rows.Cast<DataRow>()
                        .FirstOrDefault(r => r["CT_Num"]?.ToString() == ctNum);

                    string fournisseur = fournisseurRow?["CT_Intitule"]?.ToString() ?? ctNum;

                    DataTable dtArticles = new DataTable();

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        string sql = @"SELECT AR_Ref, AR_Design, StockADate, StockMinimum, StockMaximum
                           FROM VW_SEUIL_MIN_STOCK_SUM
                           WHERE CT_Num = @ctNum";

                        var ada = new SqlDataAdapter(sql, conn);
                        ada.SelectCommand.Parameters.AddWithValue("@ctNum", ctNum);
                        ada.Fill(dtArticles);
                    }

                    // ✅ FILTRAGE : articles cochés via checkStates["CT_Num|AR_Ref"]
                    DataTable dtFiltered = dtArticles.Clone();

                    foreach (DataRow row in dtArticles.Rows)
                    {
                        string arRef = row["AR_Ref"]?.ToString();
                        string articleKey = $"{ctNum}|{arRef}";

                        if (checkStates.ContainsKey(articleKey) && checkStates[articleKey])
                            dtFiltered.ImportRow(row);
                    }

                    // Aucun article coché pour ce fournisseur → on passe
                    if (dtFiltered.Rows.Count == 0)
                        continue;

                    groupedByFournisseur.Add((ctNum, fournisseur, dtFiltered));
                }

                // Sécurité globale
                if (groupedByFournisseur.Count == 0)
                {
                    MessageBox.Show("Aucun article sélectionné.", "Attention",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 3. Récapitulatif
                var recap = new StringBuilder();
                recap.AppendLine($"{groupedByFournisseur.Count} PA(s) vont être générés :\n");

                foreach (var (ctNum, fournisseur, articles) in groupedByFournisseur)
                {
                    recap.AppendLine($"📦 {fournisseur} ({articles.Rows.Count} article(s))");

                    foreach (DataRow r in articles.Rows)
                        recap.AppendLine($"   • {r["AR_Design"]}");

                    recap.AppendLine();
                }

                DialogResult confirm = MessageBox.Show(recap.ToString(), "Confirmation",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                if (confirm != DialogResult.OK) return;

                // 4. Génération
                int paGeneresCount = 0;
                var erreurs = new List<string>();
                int depot = 0, cono = 0;

                foreach (var (ctNum, fournisseur, articles) in groupedByFournisseur)
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
                                InsertFDOCENTETE(ctNum, fournisseur);
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
                            SET DO_Taxe1=0, DO_Expedit=0, DO_DateLivr=@dtt,
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
                                            cmd.Parameters.Add("@dtt", SqlDbType.DateTime).Value = dts;
                                            cmd.Parameters.Add("@dtc", SqlDbType.DateTime).Value = dtc;
                                            cmd.ExecuteNonQuery();
                                        }

                                        InsertFDOCLIGNE(dodo_piece, arRef, arDesign, qte, pu, montant, ctNum, depot, poids);
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
                        erreurs.Add($"• {fournisseur} : {ex.Message}");
                    }
                }

                // 5. Résultat final
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

                // 6. Reset sélection
                checkStates.Clear();
                UpdateSelectionCount();
                layoutView.RefreshData();

                // 6. Reset
                checkStatesFournisseurs.Clear();
                checkStatesArticles.Clear();

                UpdateSelectionCount();
                layoutView.RefreshData();
            };

            layoutView.RefreshData();
        }
    }
}