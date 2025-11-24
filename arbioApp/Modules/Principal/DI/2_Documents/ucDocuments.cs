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
//using Microsoft.Office.Interop.Outlook;
using arbioApp.Models;
using Objets100cLib;
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


        private void ucDocuments_Load(object sender, EventArgs e)
        {
            listBox1.DataSource = statutItems;
            listBox1.DisplayMember = "Text";
            listBox1.ValueMember = "Value";
            btnOuvrirDoc.Enabled = false;
            btnNouveauDoc.Enabled = false;
            btnRefresh.Enabled = false;

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

                    RepositoryItemHyperLinkEdit hyperlink = new RepositoryItemHyperLinkEdit();
                    hyperlink.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
                    hyperlink.Click += Hyperlink_Click;
                    gcEntetes.RepositoryItems.Add(hyperlink);
                    gvEntete.Columns["DO_Piece"].ColumnEdit = hyperlink;
                    gvEntete.Columns["CO_No"].VisibleIndex = -1;
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
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            string doPiece = gvEntete.GetFocusedRowCellValue("DO_Piece")?.ToString();
            OuvrirPiece(doPiece);
        }

        private void gvEntete_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {

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
            OuvrirPiece(doPiece);

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

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

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
                BarButtonItem autItem = new BarButtonItem();
                autItem.Caption= "Autorisations";
                fileMenu.ItemLinks.Add(autItem);
                autItem.ItemClick += autItem_ItemClick;
            }
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
