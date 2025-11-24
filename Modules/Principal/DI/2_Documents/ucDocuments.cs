//using Microsoft.Office.Interop.Outlook;
using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Models;
//using Objets100cLib;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Services;
using DevExpress.ChartRangeControlClient.Core;
using DevExpress.Charts.Native;
using DevExpress.CodeParser;
using DevExpress.DashboardCommon.Viewer;
using DevExpress.Utils;
using DevExpress.Xpo;
using DevExpress.XtraBars;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraExport.Helpers;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraRichEdit.Fields;
using DevExpress.XtraTreeList;
using DevExpress.XtraVerticalGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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
        public static string connectionString;
        DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit btnCloturer = new DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit();
        private readonly F_DOCENTETEService _f_DOCENTETEService;
        private readonly F_DOCENTETERepository _f_DOCENTETERepository;
        private string rec_dopiece;

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
            InitializeComponent();
            CreateDatabaseMenu();
        }


        public void SelectDoType(int doType)
        {
            listBox1.SelectedValue = doType; // Cela déclenche SelectedIndexChanged automatiquement
        }


        private void btnNouveauDoc_Click(object sender, EventArgs e)
        {
            frmMenuAchat _frmMenuAchat = new frmMenuAchat();
            _frmMenuAchat.ShowDialog();

        }

        private void BtnCloturer_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            GridView view = gcEntetes.FocusedView as GridView;
            int rowHandle = view.FocusedRowHandle;

            var recup_do_piece = view.GetRowCellValue(rowHandle, "DO_Piece");

            bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "CLOTURE");

            if (autorise)
            {
                DialogResult result = XtraMessageBox.Show(
                    "Voulez-vous vraiment clôturer le document N° " + recup_do_piece + " ?",
                    "Confirmation",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (result == DialogResult.Yes)
                {
                    using (AppDbContext context = new AppDbContext())
                    {
                        var lst = context.F_DOCENTETE.FirstOrDefault(doc => doc.DO_Piece == recup_do_piece);

                        if (lst != null)
                        {
                            lst.DO_Cloture = 1;
                            lst.DO_Type = 18;
                        }

                        context.SaveChanges();
                        MessageBox.Show("Document N °" + recup_do_piece + " clôturé avec succès!!!!", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Enabled = false;
                        ChargerDonneesDepuisBDD();
                        this.Enabled = true;
                    }
                }
            }
            else
            {
                MessageBox.Show(
                    "Vous n'avez pas l'autorisation de clôturer un document !",
                    "Clôture bloquée",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

            }
        }

        private void ucDocuments_Load(object sender, EventArgs e)
        {
            listBox1.DataSource = statutItems;
            listBox1.DisplayMember = "Text";
            listBox1.ValueMember = "Value";
            btnOuvrirDoc.Enabled = false;
            btnNouveauDoc.Enabled = false;
            btnRefresh.Enabled = false;
            gvEntete.CustomRowCellEdit += gvEntete_CustomRowCellEdit;
            btnCloturer.ButtonClick += BtnCloturer_ButtonClick;
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
            if (dbNamePrincipale == string.Empty)
            {
                return;
            }
            ChargerDonneesDepuisBDD();
        }
        public BindingSource BindingEntetes = new BindingSource();
        public void ChargerDonneesDepuisBDD()
        {

            try
            {
                if (listBox1.SelectedItem != null)
                {
                    DoTypeItem selectedItem = (DoTypeItem)listBox1.SelectedItem;
                    DoTypeSelected = selectedItem.Value;

                    Entetes.AfficherEntetes(gcEntetes, DoTypeSelected, BindingEntetes);
                    gvEntete.BestFitColumns();

                    for (int i = 0; i <= gvEntete.Columns.Count - 1; i++)
                    {
                        gvEntete.Columns[i].OptionsColumn.ReadOnly = true;
                    }
                    if (gvEntete.RowCount < 1) { return; }
                    gvEntete.Columns[0].VisibleIndex = -1;

                    GridColumn col;
                    RepositoryItemHyperLinkEdit hyperlink = new RepositoryItemHyperLinkEdit();
                    hyperlink.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
                    hyperlink.Click += Hyperlink_Click;
                    gcEntetes.RepositoryItems.Add(hyperlink);
                    gvEntete.Columns["DO_Piece"].ColumnEdit = hyperlink;
                    gvEntete.Columns["CO_No"].VisibleIndex = -1;
                    gvEntete.Columns["DO_Cours"].VisibleIndex = -1;
                    gvEntete.Columns["DO_Taxe1"].VisibleIndex = -1;
                    gvEntete.Columns["DO_CodeTaxe1"].VisibleIndex = -1;
                    gvEntete.Columns["DO_Statut"].VisibleIndex = -1;
                    gvEntete.Columns["DO_TotalTTC"].VisibleIndex = -1;
                    gvEntete.Columns["DO_Expedit"].VisibleIndex = -1;
                    gvEntete.Columns["DO_Coord01"].VisibleIndex = -1;
                    gvEntete.Columns["DE_No"].VisibleIndex = -1;

                    gvEntete.Columns["DO_TotalHT"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                    gvEntete.Columns["DO_TotalHT"].DisplayFormat.FormatString = "N2";
                    gvEntete.Columns["DO_TotalTTC"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                    gvEntete.Columns["DO_TotalTTC"].DisplayFormat.FormatString = "N2";
                    gvEntete.Columns["DO_Cours"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                    gvEntete.Columns["DO_Cours"].DisplayFormat.FormatString = "N2";

                    SetupCloturerColumn();
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gvEntete_CustomRowCellEdit(object sender, DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs e)
        {
            if (e.Column.FieldName == "Cloturer")
            {
                string doPieceValue = gvEntete.GetRowCellValue(e.RowHandle, "DO_Piece")?.ToString();

                if (!string.IsNullOrEmpty(doPieceValue) && doPieceValue.Contains("ABR"))
                {
                    if (!tester_cloturer(doPieceValue))
                    {
                        gvEntete.GridControl.RepositoryItems.Add(btnCloturer);
                        e.RepositoryItem = btnCloturer;
                    }
                    else
                    {
                        gvEntete.Columns["Cloturer"].ColumnEdit = null;
                        gvEntete.GridControl.RepositoryItems.Remove(btnCloturer);
                    }
                }
                else
                {
                    gvEntete.Columns["Cloturer"].ColumnEdit = null; 
                    gvEntete.GridControl.RepositoryItems.Remove(btnCloturer);
                }
            }
        }

        private bool tester_cloturer(string doPiece)
        {
            bool test = false;
            using (AppDbContext context = new AppDbContext())
            {
                var lst = context.F_DOCENTETE.FirstOrDefault(doc => doc.DO_Piece == doPiece);

                if (lst.DO_Cloture == 1)
                {
                    test = true;
                }
            }
            return test;
        }

        private void SetupCloturerColumn()
        {
            var oldColumn = gvEntete.Columns.ColumnByFieldName("Cloturer");
            if (oldColumn != null)
            {
                gvEntete.Columns.Remove(oldColumn);
            }

            // Retirer l'ancien RepositoryItem s’il existe déjà dans la collection
            var oldRepo = gvEntete.GridControl.RepositoryItems
                .OfType<DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit>()
                .FirstOrDefault(r => r == btnCloturer);

            if (oldRepo != null)
            {
                gvEntete.GridControl.RepositoryItems.Remove(oldRepo);
            }

            btnCloturer = new DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit();
            btnCloturer.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
            btnCloturer.Buttons.Clear();

            var cloturerButton = new DevExpress.XtraEditors.Controls.EditorButton(
                DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph)
            {
                Caption = "Clôturer",
                Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph
            };

            btnCloturer.ButtonClick += BtnCloturer_ButtonClick;

            // 🎨 Style du bouton
            cloturerButton.Appearance.BackColor = System.Drawing.Color.FromArgb(46, 204, 113); // Vert
            cloturerButton.Appearance.ForeColor = System.Drawing.Color.White;                  // Texte blanc
            cloturerButton.Appearance.Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold);
            cloturerButton.Appearance.Options.UseBackColor = true;
            cloturerButton.Appearance.Options.UseForeColor = true;
            cloturerButton.Appearance.Options.UseFont = true;
            cloturerButton.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            // Appliquer le style au bouton
            btnCloturer.Buttons.Add(cloturerButton);
            btnCloturer.ButtonsStyle = DevExpress.XtraEditors.Controls.BorderStyles.Flat; // bord plat moderne
            btnCloturer.AutoHeight = false;

            // Ajouter au grid
            gvEntete.GridControl.RepositoryItems.Add(btnCloturer);

            // Créer la colonne
            DevExpress.XtraGrid.Columns.GridColumn cloturerColumn = new DevExpress.XtraGrid.Columns.GridColumn();
            cloturerColumn.Caption = "Action finale";
            cloturerColumn.FieldName = "Cloturer";
            cloturerColumn.Visible = true;
            cloturerColumn.VisibleIndex = gvEntete.Columns.Count;
            cloturerColumn.ColumnEdit = btnCloturer;
            cloturerColumn.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowAlways;

            // Centrer le contenu dans la cellule
            cloturerColumn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            cloturerColumn.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;

            // Ajouter la colonne au GridView
            gvEntete.Columns.Add(cloturerColumn);

        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(dbNamePrincipale == string.Empty)
            {
                return;
            }
                ChargerDonneesDepuisBDD();
        }
        public void RafraichirListeDocumentsParDoType(int doType)
        {
            listBox1.SelectedValue = doType; // Cela déclenche automatiquement SelectedIndexChanged
        }

        private void Hyperlink_Click(object sender, EventArgs e)
        {
            btnOuvrirDoc_Click(sender, e);
            btnRefresh_Click(sender, e);
        }

        private void gvEntete_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName == "DO_Imprim")
            {
                if (e.Value != null)
                {
                    int value;
                    if (int.TryParse(e.Value.ToString(), out value))
                    {
                        e.DisplayText = ""; // Masquer le texte (optionnel)
                    }
                }
            }
        }

        private void gvEntete_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.Column.FieldName == "DO_Imprim") // Vérifie si c'est la colonne ciblée
            {
                int value = Convert.ToInt32(gvEntete.GetRowCellValue(e.RowHandle, "DO_Imprim"));

                if (value == 1)
                {
                    e.Appearance.Image = imageCollection1.Images[0]; // Affiche l'icône pour 1
                }
                else
                {
                    e.Appearance.Image = imageCollection1.Images[1]; // Affiche l'icône pour 0
                }
            }

        }

        private void gvEntete_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName == "DO_Imprim") // Vérifie si c'est la colonne ciblée
            {
                int value = Convert.ToInt32(gvEntete.GetRowCellValue(e.RowHandle, "DO_Imprim"));
                Image icon = value == 1 ? imageCollection1.Images[0] : imageCollection1.Images[1];

                // Calcul du centrage dans la cellule
                int iconWidth = 16;  // Largeur de l'icône
                int iconHeight = 16; // Hauteur de l'icône
                int x = e.Bounds.X + (e.Bounds.Width - iconWidth) / 2; // Centrage horizontal
                int y = e.Bounds.Y + (e.Bounds.Height - iconHeight) / 2; // Centrage vertical

                e.Graphics.DrawImage(icon, x, y, iconWidth, iconHeight);
                e.Handled = true; // Empêche le texte de s'afficher
            }
        }

        public static DateTime doDate;
        public static DateTime doDateLivrPrev;
        public static int doImprim;
        public static string doTiers;
        public static int doStatut;
        public static int doReliquat;
        public static int deno;
        public static string doRef;
        public static decimal doCours;
        public static int TypeAchat;
        public static int doExpedit;
        public static string doPiece;
        public static int doTaxe1;
        public static string doCodeTaxe1;
        public static int? CoNo;
        
        public static string doEntete;
        public static string a_type;


        private void OuvrirPiece(string dopiece)
        {
            try
            {
                if (gvEntete.FocusedRowHandle >= 0)
                {

                    //*******************************************************
                    doTiers = gvEntete.GetFocusedRowCellValue("DO_Tiers")?.ToString();
                    doRef = gvEntete.GetFocusedRowCellValue("DO_Ref")?.ToString();
                    doStatut = Convert.ToInt16(gvEntete.GetFocusedRowCellValue("DO_Statut"));
                    doDate = Convert.ToDateTime(gvEntete.GetFocusedRowCellValue("DO_Date"));
                    doCours = Convert.ToDecimal(gvEntete.GetFocusedRowCellValue("DO_Cours"));
                    doDateLivrPrev = Convert.ToDateTime(gvEntete.GetFocusedRowCellValue("DO_DateLivr"));
                    int? CO_No = Convert.ToInt16(gvEntete.GetFocusedRowCellValue("CO_No"));
                    CoNo = CO_No;

                    using (var context = new AppDbContext())
                    {
                        var conn = context.Database.Connection; // EF6
                        //MessageBox.Show($"Base = {conn.Database}, Serveur = {conn.DataSource}");
                    }


                    doCollaborateur = _collaborateurRepository.GetBy_CO_No(CO_No)                        
                    ?? throw new System.Exception($"Collaborateur introuvable pour CO_No={CO_No}");
                    

                    doEntete = gvEntete.GetFocusedRowCellValue("DO_Coord01")?.ToString();
                    deno = Convert.ToInt16(gvEntete.GetFocusedRowCellValue("DE_No"));
                    doCodeTaxe1 = gvEntete.GetFocusedRowCellValue("DO_CodeTaxe1")?.ToString();
                    doTaxe1 = Convert.ToInt16(gvEntete.GetFocusedRowCellValue("DO_Taxe1"));
                    doExpedit = Convert.ToInt16(gvEntete.GetFocusedRowCellValue("DO_Expedit"));


                    TypeAchat = Convert.ToInt32(gvEntete.GetFocusedRowCellValue("DO_Type"));
                    doPiece = gvEntete.GetFocusedRowCellValue("DO_Piece")?.ToString();
                    doImprim = Convert.ToInt16(gvEntete.GetFocusedRowCellValue("DO_Imprim"));
                    doReliquat = Convert.ToInt16(gvEntete.GetFocusedRowCellValue("DO_Reliquat"));

                    a_type = gvEntete.GetFocusedRowCellValue("A_TYPE")?.ToString();


                    if (doCollaborateur != null)
                    {
                        // Utilise doCollaborateur ici
                    }
                    else
                    {
                        MessageBox.Show("Collaborateur non trouvé.");
                    }
                    string nodoc = doPiece.Substring(3, 8);
                    string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";

                    // Vérifie si le dossier existe, sinon le créer
                    if (!Directory.Exists(destinationFolderdoc))
                    {
                        Directory.CreateDirectory(destinationFolderdoc);
                    }
                    frmEditDocument editForm = new frmEditDocument(doPiece, a_type, this, BindingEntetes);
                    editForm.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Veuillez sélectionner une ligne.", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            
        } //


        private void btnOuvrirDoc_Click(object sender, EventArgs e)
        {
            string doPiece = gvEntete.GetFocusedRowCellValue("DO_Piece")?.ToString();

            if (doPiece.ToString().StartsWith("AFA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Facture", "VIEW");

                if (autorise)
                {
                    OuvrirPiece(doPiece);
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation d'ouvrir une facture !",
                        "Ouverture bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                }
            }
            else if (doPiece.ToString().StartsWith("APA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "VIEW");

                if (autorise)
                {
                    OuvrirPiece(doPiece);
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation d'ouvrir un projet d'achat !",
                        "Ouverture bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (doPiece.ToString().StartsWith("ABC"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "VIEW");

                if (autorise)
                {
                    OuvrirPiece(doPiece);
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation d'ouvrir un bon de commande !",
                        "Ouverture bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (doPiece.ToString().StartsWith("ABR"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "VIEW");

                if (autorise)
                {
                    OuvrirPiece(doPiece);
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation d'ouvrir un Bon de réception !",
                        "Ouverture bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }

        }
        List<DoTypeItem> statutItems = new List<DoTypeItem>
            {
                new DoTypeItem { Value = 200, Text = "Tous les documents" },
                new DoTypeItem { Value = 10, Text = "Projet d'achat" },
                new DoTypeItem { Value = 12, Text = "Bon de Commande" },
                new DoTypeItem { Value = 16, Text = "Facture" },
                new DoTypeItem { Value = 13, Text = "Packing list" },
                new DoTypeItem { Value = 18, Text = "Bon de réception" },
                //new DoTypeItem { Value = 17, Text = "Facture comptabilisée" },
                //new DoTypeItem { Value = 18, Text = "Archive" },

            };

        private void CreateDatabaseMenu()
        {
           DataTable databases = GetDatabasesFromF_ACHATFILES();

            BarSubItem fileMenu = new BarSubItem();
            fileMenu.Caption = "Fichier";

            barManager1.Items.Add(fileMenu);
            barManager1.MainMenu.LinksPersistInfo.Add(new LinkPersistInfo(fileMenu));

           
            foreach (DataRow row in databases.Rows)
            {
                BarButtonItem dbItem = new BarButtonItem();
                dbItem.Caption = row["DBNAME"].ToString();

                
                dbItem.Tag = new
                {
                    ServerName = row["SERVERNAME"].ToString(),
                    ServerIP = row["SERVERIP"].ToString(),
                    ID = row["ID"].ToString()
                };

                
                dbItem.ItemClick += DbItem_ItemClick;               
                fileMenu.ItemLinks.Add(dbItem);
            }
            if (FrmMdiParent.UserRole.Contains("Administrators"))
            {
                BarSubItem autItem = new BarSubItem();
                autItem.Caption= "Autorisation";
                fileMenu.ItemLinks.Add(autItem);

                BarButtonItem autGlogale = new BarButtonItem();
                autGlogale.Caption = "Globale";
                autGlogale.ItemClick += autGlogale_ItemClick;

                BarButtonItem autAchat = new BarButtonItem();
                autAchat.Caption = "Type de documents";
                autAchat.ItemClick += autAchat_ItemClick;

                autItem.ItemLinks.Add(autGlogale);
                autItem.ItemLinks.Add(autAchat).BeginGroup = true;
            }
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
        private F_COLLABORATEURRepository _collaborateurRepository;
        F_COLLABORATEUR doCollaborateur;
        public static string dbNamePrincipale = string.Empty;
        public static string serverNamePrincipale = string.Empty;
        public static string serverIpPrincipale = string.Empty;
        private void DbItem_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                if (e.Item is BarButtonItem buttonItem)
                {
                    
                    if (buttonItem.Tag != null)
                    {                                              
                        dynamic tag = buttonItem.Tag;
                        ucDocuments.dbNamePrincipale = buttonItem.Caption;
                        ucDocuments.serverNamePrincipale = tag.ServerName;
                        ucDocuments.serverIpPrincipale = tag.ServerIP;

                        var dbContext = new AppDbContext();
                        _collaborateurRepository = new F_COLLABORATEURRepository(dbContext);

                        ChargerDonneesDepuisBDD();
                        btnOuvrirDoc.Enabled = true;
                        btnNouveauDoc.Enabled = true;
                        btnRefresh.Enabled = true;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            ChargerDonneesDepuisBDD();
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            gcEntetes.ShowPrintPreview();
        }
    }
    public class DoTypeItem
    {
        public int Value { get; set; }
        public string Text { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }
}
