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
            string arRef = layoutView.GetRowCellValue(rowHandle, "AR_Ref")?.ToString() ?? "";
            return $"{ctNum}|{arRef}";
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
                string query = "SELECT * FROM VW_SEUIL_MIN_STOCK_SUM";
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

            layoutView.CardMinSize = new Size(280, 200);

            string FormatNum(object val)
            {
                if (double.TryParse(val?.ToString(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double number))
                    return number.ToString("N2", new System.Globalization.CultureInfo("fr-FR"));
                return "0,00";
            }

            layoutView.CustomDrawCardCaption += (sender, e) =>
            {
                e.DefaultDraw();

                string key = GetRowKey(e.RowHandle);
                if (!checkStates.ContainsKey(key))
                    checkStates[key] = false;

                bool isChecked = checkStates[key]; 

                int cbSize = 16;
                int margin = 6;

                Rectangle cbRect = new Rectangle(
                    e.CaptionBounds.Right - cbSize - margin,
                    e.CaptionBounds.Top + (e.CaptionBounds.Height - cbSize) / 2,
                    cbSize,
                    cbSize);

                using (SolidBrush cbBg = new SolidBrush(isChecked
                    ? Color.FromArgb(39, 174, 96) : Color.White))
                    e.Cache.FillRectangle(cbBg, cbRect);

                using (Pen cbPen = new Pen(isChecked
                    ? Color.FromArgb(27, 130, 70)
                    : Color.FromArgb(200, 200, 200), 1.5f))
                    e.Cache.DrawRectangle(cbPen, cbRect);

                if (isChecked)
                {
                    using (Pen checkPen = new Pen(Color.White, 2f))
                    {
                        e.Cache.DrawLine(checkPen,
                            new Point(cbRect.Left + 3, cbRect.Top + cbSize / 2),
                            new Point(cbRect.Left + cbSize / 2 - 1, cbRect.Bottom - 3));
                        e.Cache.DrawLine(checkPen,
                            new Point(cbRect.Left + cbSize / 2 - 1, cbRect.Bottom - 3),
                            new Point(cbRect.Right - 2, cbRect.Top + 3));
                    }
                }
            };

            layoutView.CustomDrawCardBackground += (sender, e) =>
            {
                e.DefaultDraw();
                cardBoundsDict[e.RowHandle] = e.Bounds;

                double GetVal(string colName) => double.TryParse(
                    layoutView.GetRowCellValue(e.RowHandle, colName)?.ToString(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double v) ? v : 0;

                double rawStockADate = GetVal("StockADate");
                double rawStockMin = GetVal("StockMinimum");
                double rawStockMax = GetVal("StockMaximum");

                Color footerColor;
                if (rawStockADate < rawStockMin)
                    footerColor = Color.FromArgb(255, 220, 50, 50);
                else if (rawStockADate < rawStockMax)
                    footerColor = Color.FromArgb(255, 230, 126, 34);
                else
                    footerColor = Color.FromArgb(255, 41, 98, 162);

                int footerHeight = 55;
                Rectangle footerRect = new Rectangle(
                    e.Bounds.Left + 2,
                    e.Bounds.Bottom - footerHeight - 2,
                    e.Bounds.Width - 4,
                    footerHeight);

                using (SolidBrush bgBrush = new SolidBrush(footerColor))
                    e.Cache.FillRectangle(bgBrush, footerRect);

                using (Pen pen = new Pen(Color.White, 1))
                    e.Cache.DrawLine(pen,
                        new Point(footerRect.Left, footerRect.Top),
                        new Point(footerRect.Right, footerRect.Top));

                int lineHeight = footerHeight / 3;

                using (Font font = new Font("Tahoma", 7.5f, FontStyle.Regular))
                using (SolidBrush fg = new SolidBrush(Color.White))
                {
                    var sf = new StringFormat
                    {
                        Alignment = StringAlignment.Near,
                        LineAlignment = StringAlignment.Center
                    };

                    e.Cache.DrawString($"Stock à date : {FormatNum(rawStockADate)}", font, fg,
                        new Rectangle(footerRect.Left + 5, footerRect.Top + 2,
                                      footerRect.Width - 10, lineHeight), sf);
                    e.Cache.DrawString($"Stock Min    : {FormatNum(rawStockMin)}", font, fg,
                        new Rectangle(footerRect.Left + 5, footerRect.Top + lineHeight + 2,
                                      footerRect.Width - 10, lineHeight), sf);
                    e.Cache.DrawString($"Stock Max    : {FormatNum(rawStockMax)}", font, fg,
                        new Rectangle(footerRect.Left + 5, footerRect.Top + (lineHeight * 2) + 2,
                                      footerRect.Width - 10, lineHeight), sf);
                }
            };


            gdList.MouseClick += (sender, e) =>
            {
                Point pt = e.Location;
                LayoutViewHitInfo hi = layoutView.CalcHitInfo(pt);
                int rowHandle = hi.RowHandle;

                if (rowHandle < 0) return;
                if (!cardBoundsDict.ContainsKey(rowHandle)) return;

                Rectangle bounds = cardBoundsDict[rowHandle];
                int headerHeight = 24;

                Rectangle headerRect = new Rectangle(
                    bounds.Left, bounds.Top, bounds.Width, headerHeight);

                if (headerRect.Contains(pt))
                {
                    string key = GetRowKey(rowHandle);
                    if (!checkStates.ContainsKey(key))
                        checkStates[key] = false;

                    checkStates[key] = !checkStates[key];

                    UpdateSelectionCount();
                    layoutView.RefreshData();
                }
            };


            PanelControl searchPanel = new PanelControl();
            searchPanel.Dock = DockStyle.Top;
            searchPanel.Height = 40;
            searchPanel.BackColor = Color.FromArgb(240, 240, 240);

            LabelControl lblSearch = new LabelControl();
            lblSearch.Text = "Recherche :";
            lblSearch.Location = new Point(10, 12);
            lblSearch.Parent = searchPanel;

            TextEdit txtSearch = new TextEdit();
            txtSearch.Properties.NullValuePrompt = "🔍 Fournisseur ou article...";
            txtSearch.Properties.NullValuePromptShowForEmptyValue = true;
            txtSearch.Location = new Point(90, 8);
            txtSearch.Width = 300;
            txtSearch.Height = 24;
            txtSearch.Parent = searchPanel;

            SimpleButton btnClear = new SimpleButton();
            btnClear.Text = "✕";
            btnClear.Location = new Point(395, 8);
            btnClear.Width = 30;
            btnClear.Height = 24;
            btnClear.Parent = searchPanel;

            this.Controls.Add(searchPanel);
            searchPanel.BringToFront();
            this.Controls.Add(bottomPanel);

            void UpdateSelectionCount()
            {
                int count = checkStates.Count(kv => kv.Value);
                lblCount.Text = $"{count} article(s) sélectionné(s)";
                btnGenerer.Enabled = count > 0;
            }

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
                        dt.DefaultView.RowFilter =
                            $"CT_Intitule LIKE '%{safe}%' OR AR_Design LIKE '%{safe}%'";
                    }

                    gdList.DataSource = null;
                    gdList.DataSource = dt.DefaultView;
                    gdList.MainView = layoutView;
                    gdList.RefreshDataSource();
                    cardBoundsDict.Clear(); 
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur filtre : " + ex.Message);
                }
            };

            btnClear.Click += (sender, e) =>
            {
                txtSearch.EditValue = null;
                dt.DefaultView.RowFilter = string.Empty;
                gdList.DataSource = null;
                gdList.DataSource = dt;
                gdList.MainView = layoutView;
                gdList.RefreshDataSource();
                cardBoundsDict.Clear();
            };

            btnGenerer.Click += (sender, e) =>
            {
                var selected = checkStates
                    .Where(kv => kv.Value)
                    .Select(kv =>
                    {
                        string[] parts = kv.Key.Split('|');
                        string ctNum = parts[0];
                        string arRef = parts.Length > 1 ? parts[1] : "";

                        DataRow row = dt.Rows.Cast<DataRow>()
                            .FirstOrDefault(r =>
                                r["CT_Num"]?.ToString() == ctNum &&
                                r["AR_Ref"]?.ToString() == arRef);

                        if (row == null) return null;

                        return new
                        {
                            CT_Num = ctNum,
                            Fournisseur = row["CT_Intitule"]?.ToString(),
                            AR_Ref = arRef,
                            AR_Design = row["AR_Design"]?.ToString()
                        };
                    })
                    .Where(x => x != null)
                    .ToList();

                if (selected.Count == 0)
                {
                    MessageBox.Show("Aucun article sélectionné.", "Attention",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 2. Regrouper par fournisseur → 1 PA par groupe
                var groupedByFournisseur = selected
                    .GroupBy(s => new { s.CT_Num, s.Fournisseur })
                    .ToList();

                // 3. Récapitulatif de confirmation
                var recap = new StringBuilder();
                recap.AppendLine($"{groupedByFournisseur.Count} PA(s) vont être générés :\n");

                foreach (var group in groupedByFournisseur)
                {
                    recap.AppendLine($"📦 {group.Key.Fournisseur}  ({group.Count()} article(s))");
                    foreach (var item in group)
                        recap.AppendLine($"     • {item.AR_Design}");
                    recap.AppendLine();
                }

                DialogResult confirm = MessageBox.Show(
                    recap.ToString(), "Confirmation",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                if (confirm != DialogResult.OK) return;

                // 4. Génération : 1 en-tête + N lignes par fournisseur
                int paGeneresCount = 0;
                var erreurs = new List<string>();
                int depot = 0;
                int cono = 0;

                foreach (var group in groupedByFournisseur)
                {
                    doPiece = GetNextInvoiceNumber(_prefix);
                    try
                    {
                        using (var connection = new SqlConnection(connectionStrings))
                        {
                            connection.Open();

                            var triggers = new List<string>
                            {
                                "TG_CBINS_F_DOCENTETE",
                                "TG_INS_CPTAF_DOCENTETE",
                                "TG_INS_F_DOCENTETE"
                            };

                            try
                            {
                                foreach (var trigger in triggers)
                                {
                                    using (var cmd = new SqlCommand($"DISABLE TRIGGER {trigger} ON F_DOCENTETE", connection))
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                dodo_piece = doPiece;
                                InsertFDOCENTETE(group.Key.CT_Num, group.Key.Fournisseur);
                            }
                            finally
                            {
                                foreach (var trigger in triggers)
                                {
                                    using (var cmd = new SqlCommand($"ENABLE TRIGGER {trigger} ON F_DOCENTETE", connection))
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }

                        }
                        decimal cours=0;
                        decimal montant_total=0;
                       
                        foreach (var article in group)
                        {
                            // 1. Saisie de la quantité via formulaire DevExpress
                            int qte=0;
                            decimal pu;
                            decimal montant;
                            decimal poids;

                            using (var frmQte = new frm_saisieQuantite(article.AR_Ref, article.AR_Design, dodo_piece, article.Fournisseur))
                            {
                                if (frmQte.ShowDialog() != DialogResult.OK) continue; // Annuler → article suivant
                                qte = int.Parse(frmQte.txtQte.Text);
                                pu = decimal.Parse(frmQte.txtKg.Text);
                                cours=decimal.Parse(frmQte.txtcoursdevise.Text);
                                montant = qte * pu * cours;
                                montant_total += montant;
                                depot = frmQte.deno;
                                cono = frmQte.cono;
                                dts=frmQte.dt;
                                poids = Convert.ToDecimal(frmQte.txtPoids.Text);
                            }


                            // 2. Insertion avec la quantité saisie
                            using (var connection = new SqlConnection(connectionStrings))
                            {
                                connection.Open();
                                var triggers = new List<string>
                                {
                                    "TG_CBINS_F_DOCLIGNE",
                                    "TG_INS_CPTAF_DOCLIGNE",
                                    "TG_INS_F_DOCLIGNE"
                                };
                                try
                                {
                                    foreach (var trigger in triggers)
                                    {
                                        using (var cmd = new SqlCommand($"DISABLE TRIGGER {trigger} ON F_DOCLIGNE", connection))
                                            cmd.ExecuteNonQuery();
                                    }

                                    string sqlUpdate = "UPDATE F_DOCENTETE SET DO_Taxe1=0,DO_Expedit=0,DO_DateLivr=@dtt,DE_No=@deno,CO_No=@cono,DO_Imprim=0,DO_Statut=0,DO_Cours = @doCours,DO_TotalHTNet=@montant,DO_TotalHT=@montant,DO_TotalTTC=@montant WHERE DO_Piece = @doPiece";
                                    using (var cmdUpdate = new SqlCommand(sqlUpdate, connection))
                                    {
                                        cmdUpdate.Parameters.Add("@doCours", SqlDbType.Decimal).Value = cours; // valeur correcte pour Sage
                                        cmdUpdate.Parameters.Add("@doPiece", SqlDbType.Char, 13).Value = dodo_piece.PadRight(13);
                                        cmdUpdate.Parameters.Add("@montant", SqlDbType.Decimal).Value = montant_total;
                                        cmdUpdate.Parameters.Add("@deno", SqlDbType.Int).Value = depot;
                                        cmdUpdate.Parameters.Add("@cono", SqlDbType.Int).Value = cono;
                                        cmdUpdate.Parameters.Add("@dtt", SqlDbType.DateTime).Value = dts;
                                        cmdUpdate.ExecuteNonQuery();
                                    }
                                    InsertFDOCLIGNE(dodo_piece, article.AR_Ref, article.AR_Design, qte,pu,montant, group.Key.CT_Num,depot,poids); // ← quantité ajoutée
                                }
                                finally
                                {
                                    foreach (var trigger in triggers)
                                    {
                                        using (var cmd = new SqlCommand($"ENABLE TRIGGER {trigger} ON F_DOCLIGNE", connection))
                                            cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        }

                        paGeneresCount++;
                    }
                    catch (Exception ex)
                    {
                        erreurs.Add($"• {group.Key.Fournisseur} : {ex.Message}");
                    }
                }

                // 5. Rapport final
                if (erreurs.Count > 0)
                    MessageBox.Show(
                        $"{paGeneresCount} PA(s) générés.\n\nErreurs :\n{string.Join("\n", erreurs)}",
                        "Résultat partiel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show(
                        $"✅ {paGeneresCount} PA(s) générés avec succès.",
                        "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 6. Reset sélection
                checkStates.Clear();
                UpdateSelectionCount();
                layoutView.RefreshData();
            };

            layoutView.RefreshData();
        }
    }
}