//using Objets100cLib;
using arbioApp.DTO;
using arbioApp.Models;
using arbioApp.Models.Json;
using arbioApp.Modules.Helpers;
using arbioApp.Modules.Principal.Dashboard.CIAL;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using arbioApp.Modules.Principal.DI.Services;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Services;
using DevExpress.ChartRangeControlClient.Core;
using DevExpress.CodeParser;
using DevExpress.DashboardWin.Design;
using DevExpress.Data;
using DevExpress.DataAccess.DataFederation;
using DevExpress.DataAccess.Excel;
using DevExpress.DataAccess.Sql;
using DevExpress.DataAccess.UI.Excel;
using DevExpress.DataProcessing;
using DevExpress.DataProcessing.InMemoryDataProcessor;
using DevExpress.DataProcessing.InMemoryDataProcessor.GraphGenerator;
using DevExpress.Pdf.Xmp;
using DevExpress.Spreadsheet;
using DevExpress.UIAutomation;
using DevExpress.Utils.About;
using DevExpress.Xpo.DB.Helpers;
using DevExpress.XtraBars.Customization;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Mask;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraExport.Helpers;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraPrinting;
using DevExpress.XtraReports.UI;
using DevExpress.XtraRichEdit.Import.Doc;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraSpreadsheet.DocumentFormats.Xlsb;
using DevExpress.XtraSpreadsheet.Import.Xls;
using DevExpress.XtraTab;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Nodes;
using MailKit.Search;
using Microsoft.Office.Interop.Outlook;
using Org.BouncyCastle.Tls;
using Org.BouncyCastle.Tsp;

//using Syncfusion.Windows.Forms.Maps;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
////using Microsoft.Office.Interop.Outlook;
using System.Net;
using System.Reflection;
using System.Runtime.Remoting.Contexts;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using Exception = System.Exception;
using FieldInfo = DevExpress.DataAccess.Excel.FieldInfo;
using TextEdit = DevExpress.XtraEditors.TextEdit;


namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmEditDocument : DevExpress.XtraEditors.XtraForm
    {
        private static string dbPrincipale = ucDocuments.dbNamePrincipale;
        private static string serveripPrincipale = ucDocuments.serverIpPrincipale;
        //DECLARATIONS
        private static string connectionString = $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";

        private static string connectionString_f = $"Server={serveripPrincipale};Database=TRANSIT;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";
        private readonly AppDbContext _context;
        private List<F_COMPTET> _listeFrns;
        List<F_DOCENTETE> _listeDocs;
        List<F_COLLABORATEUR> _listeCollaborateurs;
        List<F_ARTICLE> _listeArticles;
        private F_DOCENTETE fDocenteteToModif;
        private readonly F_DOCENTETEService _f_DOCENTETEService;
        private readonly F_DOCENTETERepository _f_DOCENTETERepository;
        private readonly F_DOCLIGNEService _f_DOCLIGNEService;
        private readonly F_DOCLIGNERepository _f_DOCLIGNERepository;
        private readonly F_COMPTETRepository _f_COMPTETRepository;
        private readonly F_ARTICLERepository _f_ARTICLERepository;
        private readonly F_COMPTEARepository _f_COMPTEARepository;
        private readonly F_DOCLIGNEEMPLRepository _f_DOCLIGNEEMPLRepository;
        private readonly F_ARTSTOCKEMPLService _f_ARTSTOCKEMPLService;
        private readonly F_ARTSTOCKService _f_ARTSTOCKService;
        private readonly F_ARTSTOCKRepository _f_ARTSTOCKRepository;
        private readonly F_ARTICLEService _f_ARTICLEService;

        private readonly F_TAXERepository _f_TAXERepository;

        private readonly List<F_DEPOT> _listeDepots;

        private readonly string _typeDocument;
        private ucDocuments _ucDocuments;

        private string dopiece;

        private int StatutActuel;



        //update
        public frmEditDocument(string DoPiece, string typeDocument, ucDocuments parent, System.Windows.Forms.BindingSource source)
        {
            InitializeComponent();

            CustomLayout();
            _context = new AppDbContext();
            _listeDocs = _context.F_DOCENTETE
                    .Where(d => d.DO_Domaine == 1)
                    .ToList();
            chkActif.Checked = true;
            chkSommeil.Checked = true;
            _ucDocuments = parent;

            _f_COMPTETRepository = new F_COMPTETRepository(_context);
            InitializeChamps(DoPiece);
            dopiece = DoPiece;
            LoadCodeTaxe();
            InitializeGrid(gcLigneEdit, DoPiece);
            ChargerListeAcheteur();
            ChargerDepot();
            ChargerDevise();
            _listeArticles = _context.F_ARTICLE
                         .OrderBy(a => a.AR_Ref)   
                         .ToList();

            BindingArt.DataSource = _listeArticles;

            ChargerArticles();



            _f_DOCENTETERepository = new F_DOCENTETERepository(_context);
            _f_DOCENTETEService = new F_DOCENTETEService(_f_DOCENTETERepository);
            _f_DOCLIGNEEMPLRepository = new F_DOCLIGNEEMPLRepository(_context);
            _prefix = dopiecetxt.Text.Substring(0, 3);
            _typeDocument = typeDocument;

            _f_DOCLIGNERepository = new F_DOCLIGNERepository(_context);
            _f_DOCLIGNEService = new F_DOCLIGNEService(_context, _f_DOCLIGNERepository);

            _f_TAXERepository = new F_TAXERepository();
            _f_ARTICLERepository = new F_ARTICLERepository(_context);
            _f_ARTICLEService = new F_ARTICLEService(_f_ARTICLERepository);
            _listeCollaborateurs = _context.F_COLLABORATEUR.ToList();
            _f_ARTSTOCKRepository = new F_ARTSTOCKRepository(_context);
            _f_ARTSTOCKService = new F_ARTSTOCKService(_context, _f_ARTSTOCKRepository);
            SetupRepositoryItemLookUpEdit();
            LoadStatutLookup();

            bindingNavigator1.BindingSource = source;
            var bindingSource = source;

            LierControles(bindingSource);

            //txtSouche.DataBindings.Add("Text", source, "DO_Souche");
            dopiecetxt.DataBindings.Add("Text", source, "DO_Piece");

            if (DoPiece != "")
            {
                    LoadDocLie(DoPiece);
            }


            
            StatutActuel = (int)_listeDocs
                .Where(doc => doc.DO_Piece == DoPiece)
                .Select(doc => doc.DO_Statut)
                .FirstOrDefault();

        }

       
       

        //Nouveau
        public frmEditDocument(string DoPiece, ucDocuments parent, short statut)
        {
            try
            {
                _prefix = "APA";
                InitializeComponent();

                _context = new AppDbContext();
                _listeDocs = _context.F_DOCENTETE.ToList();

                _f_DOCENTETERepository = new F_DOCENTETERepository(_context);
                _f_DOCENTETEService = new F_DOCENTETEService(_f_DOCENTETERepository);

                F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(DoPiece, _listeDocs);

                chkActif.Checked = true;
                chkSommeil.Checked = false;

                _ucDocuments = parent;
                dateSaisie.Text = DateTime.Now.ToString("dd/MM/yyyy");
                _listeArticles = _context.F_ARTICLE
                         .OrderBy(a => a.AR_Ref)   
                         .ToList();

                BindingArt.DataSource = _listeArticles;


                ChargerArticles();

                dopiecetxt.Text = DoPiece;
                LoadExpedition();
                LoadCodeTaxe();


                lkExpedition.EditValue = 1;     //nouveau document APA


                lkDepot.EditValue = 1;//_f_DOCENTETEService.GetDepotNameByNo(43);      //dépôt par défaut DE_NO 
                lkEdCollaborateur.EditValue = 2;// doc.CO_No; par défaut

                InitializeGrid(gcLigneEdit, dopiecetxt.Text);
                //gcLigneEdit.DataSource = Lignes.CreateTableLigne();

                Lignes.cacherColonnes(gvLigneEdit);

                gvLigneEdit.Columns["Action"].OptionsColumn.ReadOnly = true;

                SetupRepositoryItemLookUpEdit();


                _listeCollaborateurs = _context.F_COLLABORATEUR.ToList();

                _f_COMPTETRepository = new F_COMPTETRepository(_context);
                ChargerListeFournisseurs();
                ChargerListeAcheteur();
                ChargerDepot();
                ChargerDevise();
                _f_DOCLIGNERepository = new F_DOCLIGNERepository(_context);
                _f_ARTICLERepository = new F_ARTICLERepository(_context);

                _f_DOCENTETERepository = new F_DOCENTETERepository(_context);
                _f_DOCENTETEService = new F_DOCENTETEService(_f_DOCENTETERepository);
                _f_COMPTEARepository = new F_COMPTEARepository(_context);
                _f_ARTICLEService = new F_ARTICLEService(_f_ARTICLERepository);
                _f_DOCLIGNEService = new F_DOCLIGNEService(_context, _f_DOCLIGNERepository);
                _f_TAXERepository = new F_TAXERepository();
                _f_DOCLIGNEEMPLRepository = new F_DOCLIGNEEMPLRepository(_context);
                _f_ARTSTOCKRepository = new F_ARTSTOCKRepository(_context);
                _f_ARTSTOCKService = new F_ARTSTOCKService(_context, _f_ARTSTOCKRepository);

                LoadStatutLookup();

                lkStatut.EditValue = 0;

                gvLigneEdit.CellValueChanging += gvLigneEdit_CellValueChanging;
                SetupGridView();

                //comboBoxAffaire.DataSource = _listePlanAnalitique.Where(p => p.N_Analytique == 1).Select(p => p.CA_Num + " - " + p.CA_Intitule).ToList();
                //comboBoxAffaire.SelectedIndex = -1;
                LoadDocLie(DoPiece);
                
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void ChargerArticles()
        {
            gcArticle.DataSource = null;
            gvArticle.Columns.Clear();

            gcArticle.DataSource = BindingArt;
            gvArticle.Columns.Clear();
            gvArticle.Columns.AddVisible("AR_Ref", "Référence");
            gvArticle.Columns.AddVisible("AR_Design", "Désignation");
            gvArticle.Columns.AddVisible("AR_Sommeil", "Sommeil");
            chkArActif.Checked = true;
            chkArSommeil.Checked = false;
            gvArticle.Columns["AR_Sommeil"].VisibleIndex = -1;
            FiltrerArticles();
            RepositoryItemHyperLinkEdit hyperlink = new RepositoryItemHyperLinkEdit();
            hyperlink.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            hyperlink.Click += Hyperlink_ClickArticle;
            gcArticle.RepositoryItems.Add(hyperlink);
            gvArticle.Columns["AR_Ref"].ColumnEdit = hyperlink;
        }
        private void LoadStatutLookup()
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Libelle", typeof(string));
            //_listeDocs = _context.F_DOCENTETE.ToList();
            F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(ucDocuments.doPiece, _listeDocs);
            if (doc != null)
            {
                int doctype = Convert.ToInt16(doc.DO_Type);
                if (doctype == 12)
                {
                    table.Rows.Add(0, "Saisie");
                    table.Rows.Add(1, "Confirmé");
                    table.Rows.Add(2, "Accepté");
                }
                else if (doctype == 13)
                {
                    table.Rows.Add(0, "Saisie");
                    table.Rows.Add(1, "Confirmé");
                    table.Rows.Add(2, "Receptionné");
                }
                else if (doctype == 16)
                {
                    table.Rows.Add(0, "Saisie");
                    table.Rows.Add(1, "Confirmé");
                    table.Rows.Add(2, "à livrer");
                }
                else if (doctype == 17)
                {
                    table.Rows.Add(0, "Saisie");
                    table.Rows.Add(1, "Confirmé");
                    table.Rows.Add(2, "Comptabilisé");
                }
                else if (doctype == 10)
                {
                    table.Rows.Add(0, "Saisie");
                    table.Rows.Add(1, "Confirmé");
                    table.Rows.Add(2, "Accepté");
                }
                else if (doctype == 18)
                {
                    table.Rows.Add(0, "reçu");
                    table.Rows.Add(1, "OK");
                }
            }
            else
            {
                table.Rows.Add(0, "Saisie");
                table.Rows.Add(1, "Confirmé");
                table.Rows.Add(2, "Accepté");
            }


            lkStatut.Properties.DataSource = table;
            lkStatut.Properties.DisplayMember = "Libelle";
            lkStatut.Properties.ValueMember = "ID";
            lkStatut.Properties.NullText = ""; // Optionnel : texte affiché si aucune valeur n'est sélectionnée
        }



        private void LierControles(System.Windows.Forms.BindingSource bindingSource)
        {
            lkEdFrns.DataBindings.Add("EditValue", bindingSource, "DO_Tiers");
            txtDoRef.DataBindings.Add("Text", bindingSource, "DO_Ref");
            lkStatut.DataBindings.Add("EditValue", bindingSource, "DO_Statut");
            dateSaisie.DataBindings.Add("Text", bindingSource, "DO_Date");
            datelivrprev.DataBindings.Add("Text", bindingSource, "DO_DateLivr");
            lkEdCollaborateur.DataBindings.Add("EditValue", bindingSource, "CO_No");
            txtCoord1.DataBindings.Add("Text", bindingSource, "DO_Coord01");
            lkDepot.DataBindings.Add("EditValue", bindingSource, "DE_No");
            lkCodeTaxe.DataBindings.Add("EditValue", bindingSource, "DO_CodeTaxe1");
            lkExpedition.DataBindings.Add("EditValue", bindingSource, "DO_Expedit");
            //lkDevise.DataBindings.Add("EditValue", bindingSource, "DO_Devise");  
        }
        private void SetupRepositoryItemLookUpEdit()
        {
            RepositoryItemLookUpEdit repoLookup = new RepositoryItemLookUpEdit();
            repoLookup.DataSource = GetArticleData(); // Récupère les articles depuis la base
            repoLookup.DisplayMember = "AR_Ref"; // Ce qui s'affiche dans la liste
            repoLookup.ValueMember = "AR_Ref"; // Ce qui est stocké dans la cellule
            repoLookup.PopulateColumns();
            gvLigneEdit.Columns["AR_Ref"].ColumnEdit = repoLookup; // Associe le repository à la colonne
            repoLookup.ImmediatePopup = true;

            RepositoryItemLookUpEdit repoLookupFrns = new RepositoryItemLookUpEdit();
            repoLookupFrns.DataSource = GetFrnsData();
            repoLookupFrns.DisplayMember = "CT_Num";
            repoLookupFrns.ValueMember = "CT_Num";
            repoLookupFrns.PopulateColumns();
            gvLigneEdit.Columns["CT_Num"].ColumnEdit = repoLookupFrns; // Associe le repository à la colonne
            repoLookupFrns.ImmediatePopup = true;

        }
        private DataTable GetArticleData()
        {
            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT AR_Ref, AR_Design FROM F_ARTICLE";
                using (SqlDataAdapter da = new SqlDataAdapter(query, conn))
                {
                    da.Fill(dt);
                }
            }
            return dt;
        }
        private DataTable GetFrnsData()
        {
            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT CT_Num, CT_Intitule FROM F_COMPTET WHERE CT_Type = 1";
                using (SqlDataAdapter da = new SqlDataAdapter(query, conn))
                {
                    da.Fill(dt);
                }
            }
            return dt;
        }
        private void InitializeChamps(string DoPiece)
        {
            doTaxe1txt.Text = ucDocuments.doTaxe1.ToString();
            dopiecetxt.Text = DoPiece;
            dateSaisie.Text = ucDocuments.doDate.ToString("dd/MM/yyyy");
            txtDoRef.Text = ucDocuments.doRef;
            txtCours.Text = ucDocuments.doCours.ToString();
            txtCours.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            txtCours.Properties.Mask.EditMask = "n2";
            txtCours.Properties.Mask.UseMaskAsDisplayFormat = true;
            datelivrprev.Text = ucDocuments.doDateLivrPrev.ToString("dd/MM/yyyy"); ;
            txtDoRef.Text = ucDocuments.doRef;
            var statutList = new List<dynamic>
                {
                    new { Value = 0, Text = "Saisie" },
                    new { Value = 1, Text = "Confirmé" },
                    new { Value = 2, Text = "Accepté" }
                };

            ChargerListeFournisseurs();
            LoadExpedition();

            lkStatut.Properties.DataSource = statutList;
            lkStatut.Properties.ValueMember = "Value";
            lkStatut.Properties.DisplayMember = "Text";

            lkEdFrns.EditValue = ucDocuments.doTiers;
            lkStatut.EditValue = ucDocuments.doStatut;
            lkEdCollaborateur.EditValue = ucDocuments.CoNo;
            lkDepot.EditValue = ucDocuments.deno;
            //lkCodeTaxe.EditValue = ucDocuments.doCodeTaxe1;

            LoadExpedition();
            lkExpedition.EditValue = ucDocuments.doExpedit;
            txtCoord1.Text = ucDocuments.doEntete;


        }
        List<F_COLLABORATEUR> _listeAcheteurs;
        List<F_DEPOT> _listeDepot;
        List<P_DEVISE> _listeDevise;
        private void ChargerListeFournisseurs()
        {
            _listeFrns = GetAllFournisseurs();

            lkEdFrns.Properties.DataSource = _listeFrns;
            lkEdFrns.Properties.ValueMember = "CT_Num"; // Clé réelle stockée
            lkEdFrns.Properties.DisplayMember = "CT_Intitule"; // Texte affiché
            lkEdFrns.Properties.PopulateColumns();
            lkEdFrns.Properties.Columns.Clear();

            lkEdFrns.Properties.Columns.Add(new LookUpColumnInfo("CT_Num", "Code", 50));
            lkEdFrns.Properties.Columns.Add(new LookUpColumnInfo("CT_Intitule", "Fournisseur"));

        }
        public List<F_COMPTET> GetAllFournisseurs()
        {
            string query = "SELECT CT_Num, CT_Intitule FROM F_COMPTET WHERE CT_Type = 1";
            List<F_COMPTET> fournisseurs = new List<F_COMPTET>();

            using (SqlConnection conn = new SqlConnection(connectionString_f))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                fournisseurs.Add(new F_COMPTET
                                {
                                    CT_Num = reader["CT_Num"].ToString(),
                                    CT_Intitule = reader["CT_Intitule"].ToString()
                                });
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return fournisseurs;
        }
        private void LoadExpedition()
        {

            string query = "SELECT cbIndice, E_Intitule FROM P_EXPEDITION WHERE E_Intitule <> ''"; // Remplacez par votre table et colonnes

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(query, conn);
                    DataTable dataTable = new DataTable();
                    dataAdapter.Fill(dataTable);

                    lkExpedition.Properties.DataSource = dataTable;
                    lkExpedition.Properties.DisplayMember = "E_Intitule";
                    lkExpedition.Properties.ValueMember = "cbIndice";

                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void ChargerListeAcheteur()
        {
            _listeAcheteurs = GetAllAcheteurs();
            
            lkEdCollaborateur.Properties.DataSource = _listeAcheteurs;
            lkEdCollaborateur.Properties.ValueMember = "CO_No"; // Clé réelle stockée
            lkEdCollaborateur.Properties.DisplayMember = "CO_Nom"; // Texte affiché
            lkEdCollaborateur.Properties.PopulateColumns();
            lkEdCollaborateur.Properties.Columns.Clear();

            lkEdCollaborateur.Properties.Columns.Add(new LookUpColumnInfo("CO_No", "CO_No", 50));
            lkEdCollaborateur.Properties.Columns.Add(new LookUpColumnInfo("CO_Nom", "Acheteur"));

        }
        private void ChargerDepot()
        {
            _listeDepot = Entetes.GetAllDepots();

            lkDepot.Properties.DataSource = _listeDepot;
            lkDepot.Properties.ValueMember = "DE_No"; // Clé réelle stockée
            lkDepot.Properties.DisplayMember = "DE_Intitule"; // Texte affiché
            lkDepot.Properties.PopulateColumns();
            lkDepot.Properties.Columns.Clear();

            lkDepot.Properties.Columns.Add(new LookUpColumnInfo("DE_No", "DE_No", 50));
            lkDepot.Properties.Columns.Add(new LookUpColumnInfo("DE_Intitule", "DEPOT"));
        }
        private void ChargerDevise()
        {
            var _listeDevise = _context.P_DEVISE.Where(d => d.D_Cours != 0).ToList();

            // ✅ Désabonner AVANT de configurer pour éviter le déclenchement prématuré
            /*lkDevise.EditValueChanged -= LkDevise_EditValueChanged;

            lkDevise.Properties.DataSource = _listeDevise;
            lkDevise.Properties.ValueMember = "cbMarq";
            lkDevise.Properties.DisplayMember = "D_Intitule";

            // Autoriser les valeurs nulles
            lkDevise.Properties.NullText = "[Sélectionner]";
            lkDevise.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.True;

            lkDevise.Properties.PopulateColumns();
            lkDevise.Properties.Columns.Clear();
            lkDevise.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo("cbMarq", "cbMarq", 50) { Visible = false });
            lkDevise.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo("D_Intitule", "DEVISE"));
            lkDevise.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo("D_Cours", "COURS") { Visible = false });

            var deviseMGA = _listeDevise.FirstOrDefault(d => d.D_Intitule == "MGA");
            if (deviseMGA != null)
            {
                lkDevise.EditValue = deviseMGA.cbMarq;
            }
            else
            {
                lkDevise.EditValue = null;
            }

            // ✅ Réabonner APRÈS avoir défini la valeur
            lkDevise.EditValueChanged += LkDevise_EditValueChanged;*/
        }

        private void LkDevise_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                // ✅ Vérifier que le contexte est initialisé
                if (_context == null)
                {
                    System.Diagnostics.Debug.WriteLine("⚠️ Context non initialisé");
                    return;
                }

                /*if (lkDevise.EditValue == null || lkDevise.EditValue == DBNull.Value)
                {
                    txtCours.EditValue = null;
                    return;
                }

                // ✅ Convertir en int de manière sécurisée
                int cbMarqValue;
                if (lkDevise.EditValue is int)
                {
                    cbMarqValue = (int)lkDevise.EditValue;
                }
                else if (int.TryParse(lkDevise.EditValue.ToString(), out cbMarqValue))
                {
                    // Conversion réussie
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ Type invalide: {lkDevise.EditValue.GetType().Name}");
                    txtCours.EditValue = null;
                    return;
                }

                var deviseSelectionnee = _context.P_DEVISE
                    .FirstOrDefault(d => d.cbMarq == cbMarqValue);

                if (deviseSelectionnee != null)
                {
                    txtCours.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;

                    if (deviseSelectionnee.D_Cours > 0)
                    {
                        string mask = deviseSelectionnee.D_Format;

                        if (!string.IsNullOrEmpty(mask) && mask.Contains("0"))
                            mask = mask.Replace("0", "#").Replace(",", "");

                        txtCours.Properties.Mask.EditMask = mask;
                        txtCours.Properties.Mask.UseMaskAsDisplayFormat = true;
                    }
                    else
                    {
                        txtCours.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
                        txtCours.Properties.Mask.EditMask = "n0";
                        txtCours.Properties.Mask.UseMaskAsDisplayFormat = true;
                    }
                    txtCours.EditValue = Convert.ToDecimal(deviseSelectionnee.D_Cours);
                }
                else
                {
                    txtCours.EditValue = null;
                }*/
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Erreur LkDevise_EditValueChanged: {ex.Message}");
                txtCours.EditValue = null;
            }
        }
        public List<F_COLLABORATEUR> GetAllAcheteurs()
        {
            List<F_COLLABORATEUR> Acheteurs = new List<F_COLLABORATEUR>();
            string query = "SELECT CO_No, CO_Nom + ' ' + CO_Prenom AS CO_Nom FROM F_COLLABORATEUR WHERE CO_Acheteur = 1";


            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Acheteurs.Add(new F_COLLABORATEUR
                                {
                                    CO_No = (int)reader["CO_No"],
                                    CO_Nom = reader["CO_Nom"].ToString()
                                });
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return Acheteurs;
        }

        DataSet ds;
        public void ChargerArtFrns()
        {
            try
            {
                string query1 = @"SELECT * FROM dbo.ACHAT_ART_FRNS";
                string connectionStringArbapp = $"Server={FrmMdiParent.DataSourceNameValueParent};" +
                                                $"Database=arbapp;User ID=Dev;Password=1234;" +
                                                $"TrustServerCertificate=True;Connection Timeout=120;";
                DataTable dtDepot =
                    arbioApp.Modules.Principal.BrowsSites.ExecuteQueryOnMultipleServers(connectionStringArbapp, query1);

                ds = new DataSet();
                dtDepot.TableName = "StockFrns";
                ds.Tables.Add(dtDepot);

                //gcDepot.DataSource = ds.Tables["StockFrns"];
                //gvDepot.PopulateColumns();

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        public void InitializeGrid(GridControl gc, string DoPiece)
        {
            Lignes.AfficherLignes(gcLigneEdit, DoPiece);
            Lignes.cacherColonnes(gvLigneEdit);

            if (gvLigneEdit.Columns["Montant Total en devise"] != null)
            {
                gvLigneEdit.Columns["Montant Total en devise"].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Custom;
                gvLigneEdit.Columns["Montant Total en devise"].SummaryItem.DisplayFormat = "{0:n2}";
            }

            if (gvLigneEdit.Columns["DL_MontantHT"] != null)
            {
                gvLigneEdit.Columns["DL_MontantHT"].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Custom;
                gvLigneEdit.Columns["DL_MontantHT"].SummaryItem.DisplayFormat = "{0:n2}";
            }

            if (gvLigneEdit.Columns["DL_MontantTTC"] != null)
            {
                gvLigneEdit.Columns["DL_MontantTTC"].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Custom;
                gvLigneEdit.Columns["DL_MontantTTC"].SummaryItem.DisplayFormat = "{0:n2}";
            }
            if (gvLigneEdit.Columns["DL_MontantRegle"] != null)
            {
                gvLigneEdit.Columns["DL_MontantRegle"].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Custom;
                gvLigneEdit.Columns["DL_MontantRegle"].SummaryItem.DisplayFormat = "{0:n2}";
            }

            gvLigneEdit.Columns["DL_Qte"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_Qte"].DisplayFormat.FormatString = "N2";
            gvLigneEdit.Columns["DL_Remise01REM_Valeur"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_Remise01REM_Valeur"].DisplayFormat.FormatString = "N2";
            gvLigneEdit.Columns["DL_PrixUnitaire"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_PrixUnitaire"].DisplayFormat.FormatString = "N2";
            gvLigneEdit.Columns["DL_Taxe1"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_Taxe1"].DisplayFormat.FormatString = "N2";
            gvLigneEdit.Columns["DL_PUDevise"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_PUDevise"].DisplayFormat.FormatString = "N2";
            gvLigneEdit.Columns["DL_Frais"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_Frais"].DisplayFormat.FormatString = "N2";

            gvLigneEdit.Columns["DL_NonLivre"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_NonLivre"].DisplayFormat.FormatString = "N0";

            gvLigneEdit.Columns["DL_MontantTTC"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_MontantTTC"].DisplayFormat.FormatString = "N2";

            gvLigneEdit.Columns["DL_MontantHT"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_MontantHT"].DisplayFormat.FormatString = "N2";

            gvLigneEdit.Columns["DL_MontantRegle"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_MontantRegle"].DisplayFormat.FormatString = "N2";

            gvLigneEdit.Columns["DL_PrixRU"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            gvLigneEdit.Columns["DL_PrixRU"].DisplayFormat.FormatString = "N2";
            gvLigneEdit.Columns["DL_PrixRU"].OptionsColumn.ReadOnly = true;
            gvLigneEdit.Columns["DL_Frais"].OptionsColumn.ReadOnly = true;
            gvLigneEdit.Columns["DL_MontantHT"].OptionsColumn.ReadOnly = true;
            gvLigneEdit.Columns["DL_MontantRegle"].OptionsColumn.ReadOnly = true;

            
            gvLigneEdit.BestFitColumns();
            gvLigneEdit.UpdateSummary();
        }
        public GridControl GridLigneEdit => gcLigneEdit;
        private void checkEdit_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit check = sender as CheckEdit;
            int row = gvLigneEdit.FocusedRowHandle;
            if (check != null)
            {
                bool isChecked = check.Checked;
                UPdateLigne(row);
            }
        }

        private void SetupGridView()
        {
            GridView view = gvLigneEdit;

            RepositoryItemHyperLinkEdit hyperlink = new RepositoryItemHyperLinkEdit();
            hyperlink.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            hyperlink.Click += Hyperlink_Click_RemoveLigne;
            gcLigneEdit.RepositoryItems.Add(hyperlink);
            view.Columns["Action"].ColumnEdit = hyperlink;

            RepositoryItemCheckEdit checkEdit = new RepositoryItemCheckEdit();
            gcLigneEdit.RepositoryItems.Add(checkEdit);
            view.Columns["Retenu"].ColumnEdit = checkEdit;
            DataColumn colRetenu = new DataColumn("Retenu", typeof(bool));
            checkEdit.CheckedChanged += checkEdit_CheckedChanged;
            colRetenu.DefaultValue = true;

            RepositoryItemHyperLinkEdit hyperlinkValidation = new RepositoryItemHyperLinkEdit();
            hyperlinkValidation.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            hyperlinkValidation.Click += Hyperlink_Click_UpdateLigne;
            gcLigneEdit.RepositoryItems.Add(hyperlinkValidation);
            view.Columns["Validation"].ColumnEdit = hyperlinkValidation;

            RepositoryItemHyperLinkEdit hlDLMontantRegle = new RepositoryItemHyperLinkEdit();
            hlDLMontantRegle.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            hlDLMontantRegle.Click += Hyperlink_EditDLMontantRegl;
            gcLigneEdit.RepositoryItems.Add(hlDLMontantRegle);
            view.Columns["DL_MontantRegle"].ColumnEdit = hlDLMontantRegle;
        }
        private void Hyperlink_Click_RemoveLigne(object sender, EventArgs e)
        {

            int rowHandle = gvLigneEdit.FocusedRowHandle;       //dans gridview


            object dlLigneObj = gvLigneEdit.GetRowCellValue(rowHandle, "DL_Ligne");   // dans la table
            int? DL_Ligne = dlLigneObj as int?;


            if (rowHandle >= 0)
            {
                DialogResult res = MessageBox.Show("Voulez-vous vraiment supprimer la ligne ?",
                    "Confirmation",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {

                    object arRefCobj = gvLigneEdit.GetRowCellValue(rowHandle, "AR_Ref");
                    string arRef = arRefCobj.ToString();
                    string doPiece = dopiecetxt.Text;
                    F_DOCLIGNE f_DOCLIGNE = _f_DOCLIGNERepository.GetF_DOCLIGNE_By_DO_Piece_AR_Ref_DL_Ligne(doPiece, arRef, (int)DL_Ligne);
                    int? DE_No = f_DOCLIGNE.DE_No;

                    object dlNoObj = gvLigneEdit.GetRowCellValue(rowHandle, "DL_No");   // dans la table
                    int? DL_No = dlNoObj as int?;

                    object dlPrixUnitaireObj = gvLigneEdit.GetRowCellValue((int)DL_No, "DL_PrixUnitaire");
                    decimal dlPrixUnitaire = 0;
                    if (dlPrixUnitaireObj != null && dlPrixUnitaireObj != DBNull.Value && dlPrixUnitaireObj.ToString() != "")
                    {
                        dlPrixUnitaire = Convert.ToDecimal(dlPrixUnitaireObj);
                    }
                    else
                    {
                        dlPrixUnitaire = 0;
                        gvLigneEdit.SetRowCellValue((int)DL_No, "DL_MontantHT", dlPrixUnitaire);
                    }

                    object dlMontantHTobj = gvLigneEdit.GetRowCellValue((int)DL_No, "DL_MontantHT");
                    decimal dlMontantHT = 0;
                    if (dlMontantHTobj != null && dlMontantHTobj != DBNull.Value && dlMontantHTobj.ToString() != "")
                    {
                        dlMontantHT = Convert.ToDecimal(dlMontantHTobj);
                    }
                    else
                    {
                        dlMontantHT = 0;
                        gvLigneEdit.SetRowCellValue((int)DL_No, "DL_MontantHT", dlMontantHT);
                    }

                    object value = gvLigneEdit.GetRowCellValue((int)DL_No, "DL_Qte");
                    int qteArt = 0;
                    if (value != null && value != DBNull.Value && value.ToString() != "")
                    {
                        qteArt = Convert.ToInt32(value);
                    }
                    else
                    {
                        qteArt = 0;
                        gvLigneEdit.SetRowCellValue((int)DL_No, "DL_Qte", qteArt);
                    }



                    // TODO: SUPPRESSION DOCLIGNE DANS LA BASE
                    //_f_DOCLIGNEEMPLRepository.DeleteF_DOCLIGNEEMPL(doPiece, (int)DL_Ligne);

                    //_f_ARTSTOCKEMPLService.UpdateArtstockEmpl(_typeDocument, doPiece, (int)DL_Ligne, arRef, qteArt, 0, DE_No);
                    _f_DOCLIGNEService.DeleteF_DOCLIGNE(doPiece, (int)DL_Ligne);
                    // TODO: SUPPRESSION DE LA LIGNE DANS GRIDVIEW
                    // MISE A JOUR DES AFFICHAGES                  
                    //MettreAJourTotalPoidsEtPrixTotalHT();
                    InitializeGrid(gcLigneEdit, doPiece);
                    gvLigneEdit.RefreshData();

                    F_DOCENTETE document = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(doPiece);

                    // Récupère toutes les lignes restantes liées à ce document
                    List<F_DOCLIGNE> lignes = _f_DOCLIGNERepository.GetAll_F_DOCLIGNE_Of_DOCENTETE(doPiece);

                    decimal totalHT = lignes.Sum(l => l.DL_MontantHT ?? 0);
                    decimal valFrais = document.DO_ValFrais ?? 0;
                    decimal txEscompte = document.DO_TxEscompte ?? 0;
                    decimal taxe1 = document.DO_Taxe1 ?? 0;

                    decimal totalHTNet = totalHT + valFrais - (totalHT * txEscompte / 100);
                    decimal totalTTC = totalHTNet + (totalHTNet * taxe1 / 100);

                    _f_DOCENTETERepository.UpdateDO_Totaux_HT_Net_TTC_Repo(doPiece, totalHT, totalHTNet, totalTTC);

                    _f_ARTSTOCKService.UpdateMontantEtQuantiteStock(_typeDocument, arRef, 0, qteArt, (int)DL_Ligne);



                }
            }
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoOptimization)]
        public void Purchase_HyperlinkClick(object sender, EventArgs e)
        {
            try
            {

                
                TreeListNode focusedNode = treeList1.FocusedNode;
                short typeDoc = (short)_f_DOCENTETEService.GetDocTypeNo(dopiecetxt.Text.Substring(0, 3));

                _listeDocs = _context.F_DOCENTETE.ToList();

             
                F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(dopiece, _listeDocs);

                string fdocentetectnum = doc.DO_Tiers;//_f_DOCENTETEService.GetDocByPiece(dopiecetxt.Text, _listeDocs).DO_Tiers;

                if (typeDoc > 13) return;

                if (focusedNode != null)
                {
                    string CT_Num = string.IsNullOrEmpty(focusedNode.GetValue("CT_Num")?.ToString())
                       ? fdocentetectnum
                       : focusedNode.GetValue("CT_Num").ToString();


                    string AR_Ref = focusedNode.GetValue("REFERENCE").ToString();
                    string DL_Design = focusedNode.GetValue("DESIGNATION").ToString();

                    string DO_Piece = dopiecetxt.Text;//focusedNode.GetValue("DO_Piece").ToString();
                    string DO_Date = dateSaisie.Text;
                    string DO_Ref = txtDoRef.Text;

                    decimal Qte_propice = Convert.ToDecimal(focusedNode.GetValue("STOCK MAXI")) - Convert.ToDecimal(focusedNode.GetValue("STOCK REEL"));
                    int remise = 0;
                    
                    //decimal pr_achat = Convert.ToDecimal(focusedNode.GetValue("AF_PrixAch"));

                    object pr_achatobj = focusedNode.GetValue("AF_PrixAch");
                    decimal pr_achat = 0;
                    if (pr_achatobj != null && pr_achatobj != DBNull.Value && pr_achatobj.ToString() != "")
                    {
                        pr_achat = Convert.ToDecimal(pr_achatobj);
                    }
                    else
                    {
                        pr_achat = 0;
                    }

                    string texte = doTaxe1txt.Text;
                    decimal tauxTVA = string.IsNullOrWhiteSpace(texte) ? 0 : Convert.ToDecimal(texte);


                    decimal prixNet = pr_achat * (1 - remise / 100);
                    decimal DL_MontantHT = prixNet * Qte_propice;
                    decimal DL_MontantTTC = DL_MontantHT * (1 + tauxTVA / 100);

                    if (AR_Ref != "")
                    {
                        DataTable dt = (DataTable)gcLigneEdit.DataSource;

                        bool exists = dt.AsEnumerable()
                        .Any(r => r.Field<string>("AR_Ref") == AR_Ref);

                        if (exists)
                        {
                            MessageBox.Show("Cette référence existe déjà dans la liste.","Message d'erreur",MessageBoxButtons.OK,MessageBoxIcon.Error);
                            return;
                        }
                        else
                        {

                            DataRow newRow = dt.NewRow();

                            newRow["DO_Domaine"] = 1; //ACHAT = 1
                            newRow["DO_Type"] = 10; // NOUVEAU DOCUMENT TSY MAINTSY 10
                            newRow["CT_Num"] = CT_Num;
                            newRow["DO_Piece"] = DO_Piece;
                            //newRow["cbDO_Piece"] = CT_Num;
                            newRow["DL_PieceBC"] = "";
                            //newRow["cbDL_PieceBC"] = CT_Num;
                            newRow["DL_PieceBL"] = "";
                            //newRow["cbDL_PieceBL"] = CT_Num;
                            newRow["DO_Date"] = DO_Date;
                            //newRow["DL_DateBC"] = CT_Num;
                            //newRow["DL_DateBL"] = CT_Num;
                            //newRow["DL_Ligne"] = CT_Num;
                            newRow["DO_Ref"] = DO_Ref;
                            //newRow["cbDO_Ref"] = CT_Num;
                            newRow["DL_TNomencl"] = 0;
                            newRow["DL_TRemPied"] = 0;
                            newRow["DL_TRemExep"] = 0;
                            newRow["AR_Ref"] = AR_Ref;
                            //newRow["cbAR_Ref"] = CT_Num;
                            newRow["DL_Design"] = DL_Design;
                            newRow["DL_Qte"] = 0;
                            newRow["DL_QteBC"] = 0;
                            newRow["DL_QteBL"] = 0;
                            //newRow["DL_PoidsNet"] = CT_Num;
                            newRow["DL_Valorise"] = 0;
                            newRow["DL_PoidsBrut"] = 0;
                            newRow["DL_Remise01REM_Valeur"] = 0;
                            newRow["DL_Remise01REM_Type"] = 0;
                            newRow["DL_Remise02REM_Valeur"] = 0;
                            newRow["DL_Remise02REM_Type"] = 0;
                            newRow["DL_Remise03REM_Valeur"] = 0;
                            newRow["DL_Remise03REM_Type"] = 0;
                            newRow["DL_PUBC"] = 0;
                            newRow["DL_PrixUnitaire"] = pr_achat;
                            newRow["DL_Taxe1"] = tauxTVA;
                            newRow["DL_TypeTaux1"] = 0;
                            newRow["DL_TypeTaxe1"] = 0;
                            newRow["DL_Taxe2"] = 0;
                            newRow["DL_TypeTaux2"] = 0;
                            newRow["DL_TypeTaxe2"] = 0;
                            newRow["CO_No"] = 0;
                            //newRow["cbCO_No"] = CT_Num;
                            newRow["AG_No1"] = 0;
                            newRow["AG_No2"] = 0;
                            newRow["DL_PrixRU"] = 0;
                            newRow["DL_CMUP"] = pr_achat;
                            newRow["DL_MvtStock"] = 0;
                            newRow["DT_No"] = 0;
                            //newRow["cbDT_No"] = CT_Num;
                            //newRow["cbAF_RefFourniss"] = CT_Num;
                            newRow["EU_Enumere"] = "";
                            newRow["EU_Qte"] = 0;
                            newRow["DL_TTC"] = 0;
                            newRow["DE_No"] = Convert.ToInt32(lkDepot.EditValue);
                            newRow["cbDE_No"] = Convert.ToInt32(lkDepot.EditValue);
                            newRow["DL_NoRef"] = 0;
                            newRow["DL_TypePL"] = 0;
                            newRow["DL_PUDevise"] = 0;
                            newRow["DL_PUTTC"] = pr_achat;
                            //newRow["DL_No"] = 0;
                            //newRow["DO_DateLivr"] = CT_Num;
                            newRow["CA_Num"] = "";
                            //newRow["cbCA_Num"] = CT_Num;
                            newRow["DL_Taxe3"] = 0;
                            newRow["DL_TypeTaux3"] = 0;
                            newRow["DL_TypeTaxe3"] = 0;
                            //newRow["cbAR_RefCompose"] = CT_Num;
                            newRow["AC_RefClient"] = "";
                            newRow["DL_MontantHT"] = DL_MontantHT;
                            newRow["DL_MontantTTC"] = DL_MontantTTC;
                            newRow["DL_FactPoids"] = 0;
                            newRow["DL_Escompte"] = 0;
                            newRow["DL_PiecePL"] = "";
                            //newRow["cbDL_PiecePL"] = CT_Num;
                            //newRow["DL_DatePL"] = CT_Num;
                            newRow["DL_QtePL"] = 0;
                            newRow["DL_NoColis"] = "";
                            newRow["DL_NoLink"] = 0;
                            //newRow["cbDL_NoLink"] = CT_Num;
                            //newRow["RP_Code"] = "";
                            //newRow["cbRP_Code"] = CT_Num;
                            newRow["DL_QteRessource"] = 0;
                            //newRow["DL_DateAvancement"] = CT_Num;
                            newRow["PF_Num"] = "";
                            newRow["DL_Frais"] = 0;
                            //newRow["cbPF_Num"] = CT_Num;
                            newRow["DL_CodeTaxe1"] = "";
                            newRow["DL_CodeTaxe2"] = "";
                            newRow["DL_CodeTaxe3"] = "";
                            newRow["DL_PieceOFProd"] = 0;
                            newRow["DL_PieceDE"] = "";
                            //newRow["cbDL_PieceDE"] = CT_Num;
                            //newRow["DL_DateDE"] = CT_Num;
                            newRow["DL_QteDE"] = 0;
                            newRow["DL_Operation"] = "";
                            newRow["DL_NoSousTotal"] = 0;
                            newRow["CA_No"] = 0;
                            //newRow["cbCA_No"] = CT_Num;
                            newRow["DO_DocType"] = 10;
                            //newRow["cbProt"] = CT_Num;
                            //newRow["cbMarq"] = CT_Num;
                            //newRow["cbCreateur"] = CT_Num;
                            //newRow["cbModification"] = CT_Num;
                            //newRow["cbReplication"] = CT_Num;
                            //newRow["cbFlag"] = CT_Num;
                            //newRow["cbCreation"] = CT_Num;
                            //newRow["cbCreationUser"] = CT_Num;
                            //newRow["cbHash"] = CT_Num;
                            //newRow["cbHashVersion"] = CT_Num;
                            //newRow["cbHashDate"] = CT_Num;
                            //newRow["cbHashOrder"] = CT_Num;          
                            newRow["Retenu"] = 1;
                            newRow["Action"] = "Remove";
                            newRow["Validation"] = "Update";
                            //newRow["Insertion"] = "Add";

                            dt.Rows.Add(newRow);
                          
                            gvLigneEdit.FocusedRowHandle = dt.Rows.IndexOf(newRow);
                            lkEdFrns.Text = CT_Num;
                            gvLigneEdit.BestFitColumns();

                            // MessageBox.Show($"Ligne cliquée:\nREFERENCE: {AR_Ref}\nFRNS: {CT_Num}");
                            AddLigne();

                            //Récupération unité
                            var recup = _context.F_ARTFOURNISS
                            .FirstOrDefault(x => x.AR_Ref == AR_Ref);

                            if (recup != null)
                            {
                                string uniteLibelle = _context.P_UNITE
                                 .Where(u => u.cbMarq == recup.AF_Unite)
                                 .Select(u => u.U_Intitule)
                                 .FirstOrDefault();

                                gvLigneEdit.CustomColumnDisplayText += (s, e) =>
                                {
                                    if (e.Column.FieldName == "Unite")
                                    {
                                        e.DisplayText = uniteLibelle ?? ""; // juste la string
                                    }
                                };

                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Cet article ne possède pas de référence", "Message d'erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        
        private void Hyperlink_EditDLMontantRegl(object sender, EventArgs e)
        {
            int row = gvLigneEdit.FocusedRowHandle;
            object dlMontantRegleobj = gvLigneEdit.GetRowCellValue(row, "DL_MontantRegle");
            decimal dlMontantRegle = 0;
            if (dlMontantRegleobj != null && dlMontantRegleobj != DBNull.Value && dlMontantRegleobj.ToString() != "")
            {
                dlMontantRegle = Convert.ToDecimal(dlMontantRegleobj);
            }
            else
            {
                dlMontantRegle = 0;
                gvLigneEdit.SetRowCellValue(row, "DL_MontantRegle", dlMontantRegle);
            }

            string dlpiece = gvLigneEdit.GetRowCellValue(row, "DO_Piece").ToString();
            string arref = gvLigneEdit.GetRowCellValue(row, "AR_Ref").ToString();
            string ctnum = gvLigneEdit.GetRowCellValue(row, "CT_Num").ToString();
            string rglibelle = dlpiece + "_" + arref + "_" + ctnum;

            frmReglementFrns frm = new frmReglementFrns(rglibelle, dlMontantRegle,  this);
            frm.ShowDialog();
            
        }
        public int GetMaxDLNo()

        {

            int maxValue = 0;

            string query = "SELECT ISNULL(MAX(DL_No), 0) FROM F_DOCLIGNE";



            using (SqlConnection conn = new SqlConnection(connectionString))

            {

                conn.Open();

                using (SqlCommand cmd = new SqlCommand(query, conn))

                {

                    object result = cmd.ExecuteScalar();

                    if (result != DBNull.Value)

                        maxValue = Convert.ToInt32(result);

                }

            }



            return maxValue;

        }
        public void UPdateLigne(int row)
        {
            object poidsNet = gvLigneEdit.GetRowCellValue(row, "DL_PoidsNet");
            object qtes = gvLigneEdit.GetRowCellValue(row, "DL_Qte");
            object prixUnitaire = gvLigneEdit.GetRowCellValue(row, "DL_PrixUnitaire");
            if (prixUnitaire != DBNull.Value && Convert.ToDecimal(prixUnitaire) > 0)
            {
                if (qtes != DBNull.Value && Convert.ToDecimal(qtes) > 0)
                {
                    if (poidsNet != DBNull.Value)
                    {
                        decimal tot1 = Convert.ToDecimal(txt_poids.Text.ToString());
                        object tot2 = gvLigneEdit.GetRowCellValue(row, "DL_PoidsNet");
                        decimal d = Convert.ToDecimal(tot2.ToString());

                        if (d > tot1)
                        {
                            MessageBox.Show("Le poids ne doit pas être supérieur au poids total FRET", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            decimal totalPoids = 0;

                            for (int i = 0; i < gvLigneEdit.RowCount; i++)
                            {
                                object values = gvLigneEdit.GetRowCellValue(i, "DL_PoidsNet");

                                if (values != null && values != DBNull.Value)
                                    totalPoids += Convert.ToDecimal(values);
                            }

                            if (totalPoids > tot1)
                            {
                                MessageBox.Show("Le total des poids ne doit pas être supérieur au poids total FRET", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                laligneamettreajour = row;
                                object value = gvLigneEdit.GetRowCellValue(row, "DL_Qte");
                                int qte = 0;
                                if (value != null && value != DBNull.Value && value.ToString() != "")
                                {
                                    qte = Convert.ToInt32(value);
                                }
                                else
                                {
                                    qte = 0;
                                    gvLigneEdit.SetRowCellValue(row, "DL_Qte", qte);
                                }
                                object frns = gvLigneEdit.GetRowCellValue(row, "CT_Num");
                                if (frns != null && frns != DBNull.Value && frns.ToString() != "")
                                {
                                }
                                else
                                {
                                    MessageBox.Show("Veuillez renseigner un fournisseur!", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                object dlremiseobj = gvLigneEdit.GetRowCellValue(row, "DL_Remise01REM_Valeur");
                                int dlremise = 0;
                                if (dlremiseobj != null && dlremiseobj != DBNull.Value && dlremiseobj.ToString() != "")
                                {
                                    dlremise = Convert.ToInt32(dlremiseobj);
                                }
                                else
                                {
                                    dlremise = 0;
                                    gvLigneEdit.SetRowCellValue(row, "DL_Remise01REM_Valeur", dlremise);
                                }

                                object dlTaxe1obj = gvLigneEdit.GetRowCellValue(row, "DL_Taxe1");
                                int dlTaxe1 = 0;
                                if (dlTaxe1obj != null && dlTaxe1obj != DBNull.Value && dlTaxe1obj.ToString() != "")
                                {
                                    dlTaxe1 = Convert.ToInt32(dlTaxe1obj);
                                }
                                else
                                {
                                    dlTaxe1 = ucDocuments.doTaxe1;
                                    gvLigneEdit.SetRowCellValue(row, "DO_Taxe1", dlTaxe1);
                                }

                                object dlPuDeviseobj = gvLigneEdit.GetRowCellValue(row, "DL_PuDevise");
                                decimal dlPuDevise = 0;
                                if (dlPuDeviseobj != null && dlPuDeviseobj != DBNull.Value && dlPuDeviseobj.ToString() != "")
                                {
                                    dlPuDevise = Convert.ToDecimal(dlPuDeviseobj);
                                }
                                else
                                {
                                    dlPuDevise = 0;
                                    gvLigneEdit.SetRowCellValue(row, "DL_PuDevise", dlPuDevise);
                                }

                                object dlFraisobj = gvLigneEdit.GetRowCellValue(row, "DL_Frais");
                                decimal dlFrais = 0;
                                if (dlFraisobj != null && dlFraisobj != DBNull.Value && dlFraisobj.ToString() != "")
                                {
                                    dlFrais = Convert.ToDecimal(dlFraisobj);
                                }
                                else
                                {
                                    dlFrais = 0;
                                    gvLigneEdit.SetRowCellValue(row, "DL_Frais", dlFrais);
                                }

                                object dlNonLivreobj = gvLigneEdit.GetRowCellValue(row, "DL_NonLivr");
                                int dlNonLivre = 0;
                                if (dlNonLivreobj != null && dlNonLivreobj != DBNull.Value && dlNonLivreobj.ToString() != "")
                                {
                                    dlNonLivre = Convert.ToInt32(dlNonLivreobj);
                                }
                                else
                                {
                                    dlNonLivre = 0;
                                    gvLigneEdit.SetRowCellValue(row, "DL_NonLivre", dlNonLivre);
                                }

                                object dlMontantHTobj = gvLigneEdit.GetRowCellValue(row, "DL_MontantHT");
                                decimal dlMontantHT = 0;
                                if (dlMontantHTobj != null && dlMontantHTobj != DBNull.Value && dlMontantHTobj.ToString() != "")
                                {
                                    dlMontantHT = Convert.ToDecimal(dlMontantHTobj);
                                }
                                else
                                {
                                    dlMontantHT = 0;
                                    gvLigneEdit.SetRowCellValue(row, "DL_MontantHT", dlMontantHT);
                                }

                                object dlMontantTTCobj = gvLigneEdit.GetRowCellValue(row, "DL_MontantTTC");
                                decimal dlMontantTTC = 0;
                                if (dlMontantTTCobj != null && dlMontantTTCobj != DBNull.Value && dlMontantTTCobj.ToString() != "")
                                {
                                    dlMontantTTC = Convert.ToDecimal(dlMontantTTCobj);
                                }
                                else
                                {
                                    dlMontantTTC = 0;
                                    gvLigneEdit.SetRowCellValue(row, "DL_MontantTTC", dlMontantTTC);
                                }

                                object dlPiecefrnsobj = gvLigneEdit.GetRowCellValue(row, "DL_PieceFourniss");
                                string dlPiecefrns = dlPiecefrnsobj.ToString();
                                if (dlPiecefrnsobj != null && dlPiecefrnsobj != DBNull.Value && dlPiecefrnsobj.ToString() != "")
                                {
                                    dlPiecefrns = Convert.ToString(dlPiecefrnsobj);
                                }
                                else
                                {
                                    dlPiecefrns = string.Empty;
                                }

                                object dlDatePiecefrnsobj = gvLigneEdit.GetRowCellValue(row, "DL_DatePieceFourniss");
                                DateTime dlDatePiecefrns;

                                if (dlDatePiecefrnsobj != null && !Convert.IsDBNull(dlDatePiecefrnsobj) &&
                                    !string.IsNullOrEmpty(dlDatePiecefrnsobj.ToString()))
                                {
                                    dlDatePiecefrns = Convert.ToDateTime(dlDatePiecefrnsobj);
                                }
                                else
                                {
                                    dlDatePiecefrns = DateTime.Now;
                                }



                                int? dl_No = 0;
                                try
                                {

                                    if (row >= 0)
                                    {
                                        decimal remisePourcent = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_Remise01REM_Valeur"));
                                        decimal puBrut = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_PrixUnitaire"));
                                        decimal puNet = puBrut * (1 - remisePourcent / 100);
                                        int quantiteEcriteStock = qte;
                                        dl_No = Convert.ToInt32(gvLigneEdit.GetRowCellValue(row, "DL_No"));

                                        DataTable dt = (DataTable)gvLigneEdit.GridControl.DataSource;
                                        string doPiece = dopiecetxt.Text;//dt.Rows[row]["DO_Piece"].ToString();
                                        decimal totalHTNet = dt.AsEnumerable()
                                        .Where(r =>
                                            r["DL_MontantHT"] != DBNull.Value &&
                                            r["DO_Piece"] != DBNull.Value &&
                                            r["Retenu"] != DBNull.Value &&
                                            Convert.ToInt32(r["Retenu"]) != 0 &&
                                            r["DO_Piece"].ToString() == doPiece
                                        )
                                        .Sum(r => Convert.ToDecimal(r["DL_MontantHT"]));

                                        decimal totalTTCNet = dt.AsEnumerable()
                                                        .Where(r =>
                                                            r["DL_MontantTTC"] != DBNull.Value &&
                                                            r["DO_Piece"] != DBNull.Value &&
                                                            r["Retenu"] != DBNull.Value &&
                                                            Convert.ToInt32(r["Retenu"]) != 0 &&
                                                            r["DO_Piece"].ToString() == doPiece
                                                        )
                                                        .Sum(r => Convert.ToDecimal(r["DL_MontantTTC"]));
                                        //// Mise à jour entête de document (F_DOCENTETE)
                                        int retenu = Convert.ToInt16(gvLigneEdit.GetRowCellValue(row, "Retenu"));
                                        _f_DOCENTETEService.UpdateDO_Totaux_HT_Net_TTC(dopiecetxt.Text, puNet, quantiteEcriteStock, totalHTNet, totalTTCNet);


                                        //// Mise à jour du stock des articles (F_ARTSTOCK) (Montant du stock et quantité en stock)
                                        //int? DE_No = Convert.ToInt32(lkDepot.EditValue);
                                        //_f_ARTSTOCKService.UpdateMontantEtQuantiteStock(_typeDocument, arRef, quantiteEcriteStock, previousQuantiteEcriteStock, DE_No);



                                        //// Mise à jour du ligne de document (F_DOCLIGNE) (Mise à jour des quantités, des poids et des prix (PrixRU et CMUP))
                                        string CtNum = Convert.ToString(gvLigneEdit.GetRowCellValue(row, "CT_Num"));
                                        string? dl_Designe = Convert.ToString(gvLigneEdit.GetRowCellValue(row, "DL_Design"));
                                        string arRef = Convert.ToString(gvLigneEdit.GetRowCellValue(row, "AR_Ref"));


                                        object dlLigneObj = gvLigneEdit.GetRowCellValue(row, "DL_Ligne");
                                        int dlLigne = 0;
                                        if (dlLigneObj != null && dlLigneObj != DBNull.Value && dlLigneObj.ToString() != "")
                                        {
                                            dlLigne = (int)Convert.ToInt64(dlLigneObj);
                                        }
                                        else
                                        {
                                            dlLigne = 0;
                                            gvLigneEdit.SetRowCellValue(row, "DL_Ligne", dlLigne);
                                        }
                                        string strDL_No = Convert.ToString(gvLigneEdit.GetRowCellValue(row, "DL_No"));
                                        decimal montantRegl = _f_DOCLIGNERepository.GetMontantRegleByPieceArRef(dopiece, arRef, CtNum);
                                        decimal poids = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_PoidsNet"));
                                        var val = gvLigneEdit.GetListSourceRowCellValue(row, "Unite");

                                        string unite = val == null ? "" : val.ToString();


                                        _f_DOCLIGNEService.UpdateF_DOCLIGNE(dopiecetxt.Text, CtNum, arRef, dl_Designe, puBrut, dlLigne, quantiteEcriteStock, _typeDocument, dlTaxe1, dlMontantHT, dlMontantTTC, retenu, remisePourcent, dlPiecefrns, dlDatePiecefrns, montantRegl, poids, unite);

                                        //// Mise à jour du stock de l'article dans un emplacement concerné
                                        //_f_ARTSTOCKEMPLService.UpdateArtstockEmpl(_typeDocument, ct_Num, dl_Ligne, arRef, previousQuantiteEcriteStock, quantiteEcriteStock, DE_No);

                                        //// Mise à jour de la quantité prise dans l'emplacement concerné (DL_Qte)
                                        //_f_DOCLIGNEEMPLRepository.UpdateDL_Qte(_typeDocument, _currentDocPieceNo, dl_Ligne, quantiteEcriteStock);

                                    }
                                    InitializeGrid(gcLigneEdit, dopiecetxt.Text);
                                    gvLigneEdit.UpdateSummary();

                                }
                                catch (System.Exception ex)
                                {
                                    MethodBase m = MethodBase.GetCurrentMethod();
                                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Le poids ne doit pas être vide", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("La quantité ne peut pas être vide", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Le prix unitaire ne peut pas être vide", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //public int GetMaxValueEF()
        //{
        //    using (var ctx = new AppDbContext())
        //    {
        //        return (int)ctx.F_DOCLIGNE.Max(u => u.DL_No);
        //    }
        //}


        public int laligneamettreajour;

        private void Hyperlink_Click_UpdateLigne(object sender, EventArgs e)
        {

            int row = gvLigneEdit.FocusedRowHandle;
            laligneamettreajour = row;
            if (txtCours.Text == "" || txtCours.Text == "0,00")
            {
                MessageBox.Show("Le cours de devise est vide", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                UPdateLigne(row);
                string _currentDocPieceNo = dopiecetxt.Text;
                var entete = _context.F_DOCENTETE.FirstOrDefault(d => d.DO_Piece == _currentDocPieceNo);

                if (entete != null) {
                    entete.DO_Cours = Convert.ToDecimal(txtCours.Text.ToString());
                    _context.SaveChanges();
                }

                try
                {
                    using (AppDbContext context = new AppDbContext())
                    {
                        // Pour un AJOUT (vérifier que l'enregistrement n'existe pas déjà)
                        var existingFRET = context.F_FRETS
                            .FirstOrDefault(s => s.DO_PIECE == dopiecetxt.Text);

                        if (existingFRET == null)
                        {
                            try
                            {
                                F_FRET f = new F_FRET();
                                f.DO_PIECE = dopiecetxt.Text;
                                f.DO_MONTANT = Convert.ToDecimal(txt_prix.Text);
                                f.DO_POIDS = Convert.ToDecimal(txt_poids.Text);

                                context.F_FRETS.Add(f);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }
                        else
                        {
                            existingFRET.DO_MONTANT = Convert.ToDecimal(txt_prix.Text);
                            existingFRET.DO_POIDS = Convert.ToDecimal(txt_poids.Text);
                        }

                        context.SaveChanges();
                    }
                }
                catch (System.Data.Entity.Infrastructure.DbUpdateException ex)
                {
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erreur");
                }

                string sql = "EXEC SP_CalculCoutRevientParValeur @DO_Piece, @prix_total,@poids_total";

                _context.Database.ExecuteSqlCommand(
                   sql,
                   new SqlParameter("@DO_Piece", dopiecetxt.Text),
                   new SqlParameter("@prix_total", Convert.ToDecimal(txt_prix.Text)),
                   new SqlParameter("@poids_total", Convert.ToDecimal(txt_poids.Text))
               );

                btnEditLigne_Click(sender, e);
            }
        }
        public static void UpdateSequence(string prefix, int currentNumber)
        {
            int year = DateTime.Now.Year;
            string query = "SELECT CurrentNumber FROM ARB_ACHAT_DOPIECE WHERE Prefix = @Prefix AND Year = @Year";

            // Récupérer le dernier numéro
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    currentNumber++;
                    string updateQuery = "UPDATE ARB_ACHAT_DOPIECE SET CurrentNumber = @CurrentNumber WHERE Prefix = @Prefix AND Year = @Year";
                    using (var updateCommand = new SqlCommand(updateQuery, connection))
                    {
                        updateCommand.Parameters.AddWithValue("@CurrentNumber", currentNumber);
                        updateCommand.Parameters.AddWithValue("@Prefix", prefix);
                        updateCommand.Parameters.AddWithValue("@Year", year);
                        updateCommand.ExecuteNonQuery();
                    }
                }
            }
        }
        private void ImprimerDocument()
        {
           

            rptDocument report = new rptDocument();
            report.Parameters["pDocumentNumber"].Value = dopiecetxt.Text;

            ShowPreviewWithTracking(report, dopiecetxt.Text);
        }
        private void AddLigne()
        {
            try
            {
                int row = gvLigneEdit.FocusedRowHandle;
                if (row >= 0)
                {
                    string arRef = gvLigneEdit.GetRowCellValue(row, "AR_Ref").ToString();
                    string CT_NumClient = gvLigneEdit.GetRowCellValue(row, "CT_Num").ToString();
                    if (CT_NumClient == "")
                    {
                        MessageBox.Show("Fournisseur ?");
                        return;
                    }
                    object value = gvLigneEdit.GetRowCellValue(row, "DL_Qte");
                    decimal qte = 0;
                    if (value != null && value != DBNull.Value && value.ToString() != "")
                    {
                        qte = Convert.ToDecimal(value);
                    }
                    else
                    {
                        qte = 0;
                    }
                    //decimal DLTaxe1 = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_Taxe1"));
                    object taxe1value = gvLigneEdit.GetRowCellValue(row, "DL_Taxe1");
                    decimal DLTaxe1 = 0;

                    if (taxe1value != null && decimal.TryParse(taxe1value.ToString(), out var result))
                    {
                        DLTaxe1 = result;
                    }


                    // Retrieve the F_ARTICLE object corresponding to the AR_Ref
                    F_ARTICLE articleChoisi = _f_ARTICLERepository.GetF_ARTICLEByAR_Ref(arRef);
                    if (articleChoisi == null)
                    {
                        MessageBox.Show($"L'article avec la référence '{arRef}' est introuvable.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    int? DE_No = Convert.ToInt32(lkDepot.EditValue);
                    int? dl_Ligne = 0;


                    decimal previousQuantite = 0;

                    //decimal remisePourcent = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_Remise01REM_Valeur"));
                    object remisePourcentvalue = gvLigneEdit.GetRowCellValue(row, "DL_Remise01REM_Valeur");
                    decimal remisePourcent = 0;

                    if (remisePourcentvalue != null && decimal.TryParse(remisePourcentvalue.ToString(), out var remisePourcentresult))
                    {
                        remisePourcent = remisePourcentresult;
                    }



                    string arDesign = gvLigneEdit.GetRowCellValue(row, "DL_Design").ToString();
                    decimal puBrut = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_PrixUnitaire"));
                    decimal puNet = puBrut * (1 - remisePourcent / 100);

                    //decimal montantHT = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_MontantHT"));
                    object montantHTvalue = gvLigneEdit.GetRowCellValue(row, "DL_MontantHT");
                    decimal montantHT = 0;
                    if (montantHT == 0)
                    {
                        montantHT = puNet * qte;
                        gvLigneEdit.SetRowCellValue(row, "DL_MontantHT", montantHT);
                    }
                    //decimal montantTTC = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(row, "DL_MontantTTC"));
                    object montantTTCvalue = gvLigneEdit.GetRowCellValue(row, "DL_MontantTTC");
                    decimal montantTTC = 0;

                    if (montantTTC == 0)
                    {
                        montantTTC = montantHT * (1 + DLTaxe1 / 100);
                        gvLigneEdit.SetRowCellValue(row, "DL_MontantTTC", montantTTC);
                    }

                    int retenu = Convert.ToInt32(gvLigneEdit.GetRowCellValue(row, "Retenu"));


                    F_DOCENTETE docEnCours = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(dopiecetxt.Text);


                    F_COLLABORATEUR collab = _listeCollaborateurs.Where(c => c.CO_No == (int)lkEdCollaborateur.EditValue).FirstOrDefault();

                    DateTime DO_Date = dateSaisie.DateTime;
                    DateTime DO_DateLivr;
                    if (datelivrprev.EditValue == null || !(datelivrprev.EditValue is DateTime))
                    {
                        DO_DateLivr = new DateTime(1753, 01, 01);
                    }
                    else
                    {
                        DO_DateLivr = (DateTime)datelivrprev.EditValue;
                    }


                    DateTime dateLivrReal;
                    if (datelivrprev.EditValue == null || !(datelivrprev.EditValue is DateTime))
                    {
                        dateLivrReal = new DateTime(1753, 01, 01);
                    }
                    else
                    {
                        dateLivrReal = (DateTime)datelivrprev.EditValue;
                    }


                    if (row >= 0)
                    {
                        int? maxValueDL_Ligne = 0;

                        for (int i = 0; i < gvLigneEdit.RowCount; i++)
                        {
                            object dlligne = gvLigneEdit.GetRowCellValue(i, "DL_Ligne");
                            if (dlligne != null && int.TryParse(dlligne.ToString(), out int val))
                            {
                                if (val > maxValueDL_Ligne)
                                {
                                    maxValueDL_Ligne = val;
                                }
                            }
                        }

                        int? numeroLigneDL_Ligne = maxValueDL_Ligne + 1000;

                        string reference = txtDoRef.Text;
                        short typeDoc = (short)_f_DOCENTETEService.GetDocTypeNo(_prefix);
                        short DL_NoRef = (short)(gvLigneEdit.RowCount + 1);
                        int dlno = GetMaxDLNo() + 1;


                            _f_DOCLIGNEService.AjouterF_DOCLIGNE(
                                typeDoc,
                                CT_NumClient,
                                dopiecetxt.Text,
                                DO_Date,
                                numeroLigneDL_Ligne,
                                docEnCours,
                                arRef,
                                arDesign,
                                DLTaxe1,
                                qte,
                                typeDoc.ToString(),
                                articleChoisi,
                                qte.ToString(),
                                remisePourcent.ToString(),
                                puNet.ToString(),
                                collab,
                                DL_NoRef,
                                puBrut,
                                DO_DateLivr,
                                comboBoxAffaire.Text,
                                montantTTC.ToString(),
                                montantHT.ToString(),
                                DateTime.Now,
                                DE_No,
                                dlno,
                                retenu);


                            //UPDATE LES PRIX DANS F_DOCENTETE

                            DataTable dt = (DataTable)gvLigneEdit.GridControl.DataSource;
                            string doPiece = dopiecetxt.Text;//dt.Rows[row]["DO_Piece"].ToString();
                            decimal totalHTNet = dt.AsEnumerable()
                                            .Where(row =>
                                                row["DL_MontantHT"] != DBNull.Value &&
                                                row["DO_Piece"] != DBNull.Value &&
                                                row["Retenu"] != DBNull.Value &&
                                                Convert.ToInt32(row["Retenu"]) != 0 &&
                                                row["DO_Piece"].ToString() == doPiece
                                            )
                                            .Sum(row => Convert.ToDecimal(row["DL_MontantHT"]));
                            decimal totalTTCNet = dt.AsEnumerable()
                                            .Where(row =>
                                                row["DL_MontantTTC"] != DBNull.Value &&
                                                row["DO_Piece"] != DBNull.Value &&
                                                row["Retenu"] != DBNull.Value &&
                                                Convert.ToInt32(row["Retenu"]) != 0 &&
                                                row["DO_Piece"].ToString() == doPiece
                                            )
                                            .Sum(row => Convert.ToDecimal(row["DL_MontantTTC"]));
                            _f_DOCENTETEService.UpdateDO_Totaux_HT_Net_TTC(dopiecetxt.Text, puNet, 0, totalHTNet, totalTTCNet);

                            // Mise à jour F_ARTSTOCK
                            object rawValue = gvLigneEdit.GetRowCellValue(row, "DL_Qte");

                            decimal quantiteEcriteStock = 0;

                            if (rawValue != null && !string.IsNullOrWhiteSpace(rawValue.ToString()))
                            {
                                decimal.TryParse(rawValue.ToString(), out quantiteEcriteStock);
                            }
                            else
                            {
                                quantiteEcriteStock = 0; // ou une valeur par défaut, ou lever une alerte si c'est anormal
                            }
                            //decimal quantiteEcriteStock = int.Parse(gvLigneEdit.GetRowCellValue(row, "DL_Qte").ToString());
                            _f_ARTSTOCKService.UpdateMontantEtQuantiteStock(_typeDocument, arRef, quantiteEcriteStock, previousQuantite, DE_No);

                            // Mise à jour F_ARTSTOCKEMPL
                            //_f_ARTSTOCKEMPLService.UpdateArtstockEmpl(_typeDocument, CT_NumClient, numeroLigneDL_Ligne, arRef, previousQuantite, quantiteEcriteStock, DE_No);
                    }
                }

                InitializeGrid(gcLigneEdit, dopiecetxt.Text);
                gvLigneEdit.UpdateSummary();
                gvLigneEdit.RefreshData();
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool not_existence_ref(string cond)
        {

            bool exist = true;
            using (SqlConnection connection = new SqlConnection(
                $"Data Source=SRV-ARB;Initial Catalog=ARBIOCHEM;User ID=Dev;Password=1234;TrustServerCertificate=True"))
                {
                    try
                    {
                        connection.Open();

                        // Vérifier si l'utilisateur a changé son mot de passe
                        string query = "SELECT AR_Ref FROM dbo.F_ARTICLE WHERE AR_Ref = @ar_ref";
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@ar_ref", cond);

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    exist = false;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MethodBase m = MethodBase.GetCurrentMethod();
                        MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            return exist;
        }
        private void Hyperlink_Click_AddLigne(object sender, EventArgs e)
        {
            AddLigne();
        }

        private DataTable ExecuteQuery(string connectionString, string query)
        {
            DataTable dt = new DataTable();

            try
            {
                // Modifier le timeout pour éviter les blocages
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connectionString);
                builder.ConnectTimeout = 5;
                builder.ConnectRetryCount = 0;

                using (SqlConnection conn = new SqlConnection(builder.ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 10;
                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur SQL : {ex.Message}");
                // Retourner une DataTable vide sans afficher d'erreur
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur : {ex.Message}");
            }

            return dt;
        }


        public void ExecuteStockAlert()
        {
            treeList1.BeginUpdate();

            try
            {
                // Supprime tous les nœuds
                treeList1.ClearNodes();

                // Supprime toutes les colonnes
                treeList1.Columns.Clear();
            }
            finally
            {
                // Débloque la mise à jour et rafraîchit l'affichage
                treeList1.EndUpdate();
            }

            try
            {
                string query1 = @"SELECT * FROM dbo.VW_ETAT_STOCK WHERE CT_INTITULE='" + lkEdFrns.Text + "' AND DEPOT IN (" + labelControl8.Text + ")";


                string connectionStringArbapp = $"Server={FrmMdiParent.DataSourceNameValueParent};" +
                                 $"Database=TRANSIT;" +
                                 $"User ID=Dev;" +
                                 $"Password=1234;" +
                                 $"TrustServerCertificate=True;" +
                                 $"Connection Timeout=120;";

                DataTable dtMaster = null;
                try
                {
                    if (string.IsNullOrEmpty(connectionStringArbapp))
                    {
                        // Log l'erreur
                        throw new ArgumentException("Chaîne de connexion vide");
                    }

                    // Utilisation
                    dtMaster = ExecuteQuery(connectionStringArbapp, query1);
                }
                catch (SqlException sqlEx)
                {
                    // Log l'erreur SQL spécifique
                    System.Diagnostics.EventLog.WriteEntry("Application",
                        $"Erreur SQL: {sqlEx.Message}\nCode: {sqlEx.Number}\nServeur: {sqlEx.Server}",
                        System.Diagnostics.EventLogEntryType.Error);
                    return;
                }
                catch (TimeoutException timeEx)
                {
                    // Log timeout
                    System.Diagnostics.EventLog.WriteEntry("Application",
                        $"Timeout: {timeEx.Message}",
                        System.Diagnostics.EventLogEntryType.Warning);
                    return;
                }
                catch (Exception ex)
                {
                    // Log toute autre erreur
                    System.Diagnostics.EventLog.WriteEntry("Application",
                        $"Erreur: {ex.Message}\nStack: {ex.StackTrace}",
                        System.Diagnostics.EventLogEntryType.Error);
                    return;
                }

                if (dtMaster == null || dtMaster.Rows.Count == 0)
                {
                    return;
                }

                if (dtMaster == null)
                    throw new InvalidOperationException("La requête a réussi, mais a retourné une table nulle.");

                // Continuer le traitement...

                treeList1.BeginUpdate();
                treeList1.ClearNodes();
                treeList1.Columns.Clear();

                var parentColumns = new[]
                    { "SITE", "FAMILLE", "REFERENCE", "DESIGNATION", "CT_Num", "CT_Intitule", "PURCHASE", "AF_PrixAch" };
                var childColumns = new[] { "DEPOT", "STOCK REEL", "STOCK MINI", "STOCK MAXI" };

                foreach (string col in parentColumns)
                {
                    treeList1.Columns.AddVisible(col);
                }

                foreach (string col in childColumns)
                {
                    treeList1.Columns.AddVisible(col);
                }

                var groupedData = dtMaster.AsEnumerable()
                    .GroupBy(r => new
                    {
                        Site = r["SITE"],
                        Famille = r["FAMILLE"],
                        Reference = r["REFERENCE"],
                        Designation = r["DESIGNATION"],
                        CtNum = r["CT_Num"],
                        CtIntitule = r["CT_Intitule"],
                        Purchase = r["PURCHASE"],
                        AF_PrixAch = r["AF_PrixAch"]
                    });

                if (groupedData.Any())
                {
                    foreach (var group in groupedData)
                    {
                        decimal totalStockReel = group.Sum(r => r.Field<decimal?>("STOCK REEL") ?? 0);
                        decimal totalStockMini = group.Sum(r => r.Field<decimal?>("STOCK MINI") ?? 0);
                        decimal totalStockMaxi = group.Sum(r => r.Field<decimal?>("STOCK MAXI") ?? 0);
                        TreeListNode parentNode = treeList1.AppendNode(new object[]
                        {
                    group.Key.Site,
                    group.Key.Famille,
                    group.Key.Reference,
                    group.Key.Designation,
                    group.Key.CtNum,
                    group.Key.CtIntitule,
                    group.Key.Purchase,
                    group.Key.AF_PrixAch,
                    "Tous", // DEPOT
                    totalStockReel, // STOCK REEL total
                    totalStockMini, // STOCK MINI total
                    totalStockMaxi // STOCK MAXI total
                        }, null);

                        foreach (DataRow childRow in group)
                        {
                            treeList1.AppendNode(new object[]
                            {
                        null, // SITE
                        null, // FAMILLE
                        null, // REFERENCE
                        null, // DESIGNATION
                        null, // CT_Num
                        null, // CT_Intitule
                        null, // PURCHASE
                        null, // AF_PrixAch
                        childRow["DEPOT"],
                        childRow["STOCK REEL"],
                        childRow["STOCK MINI"],
                        childRow["STOCK MAXI"]
                            }, parentNode);
                        }
                    }
                }

                RepositoryItemHyperLinkEdit repo = new RepositoryItemHyperLinkEdit();
                treeList1.RepositoryItems.Add(repo);
                treeList1.Columns["PURCHASE"].ColumnEdit = repo;
                repo.Click += Purchase_HyperlinkClick;

                treeList1.Columns["STOCK REEL"].Format.FormatType = DevExpress.Utils.FormatType.Numeric;
                treeList1.Columns["STOCK REEL"].Format.FormatString = "N2";
                treeList1.Columns["STOCK MINI"].Format.FormatType = DevExpress.Utils.FormatType.Numeric;
                treeList1.Columns["STOCK MINI"].Format.FormatString = "N2";
                treeList1.Columns["STOCK MAXI"].Format.FormatType = DevExpress.Utils.FormatType.Numeric;
                treeList1.Columns["STOCK MAXI"].Format.FormatString = "N2";

                treeList1.NodeCellStyle += TreeList1_NodeCellStyle;

                treeList1.EndUpdate();
            }
            catch (Exception ex) 
            {
                return;
            }
            
        }

        private void TreeList1_NodeCellStyle(object sender, GetCustomNodeCellStyleEventArgs e)
        {
            if (e.Node.ParentNode == null) // Uniquement pour les nœuds parents
            {
                if (e.Column.FieldName == "PURCHASE")
                    return;
                // Récupérer les valeurs des colonnes
                object stockReelObj = e.Node.GetValue("STOCK REEL");
                object stockMiniObj = e.Node.GetValue("STOCK MINI");

                if (stockReelObj != null && stockMiniObj != null &&
                    stockReelObj != DBNull.Value && stockMiniObj != DBNull.Value)
                {
                    decimal stockReel = Convert.ToDecimal(stockReelObj);
                    decimal stockMini = Convert.ToDecimal(stockMiniObj);

                    if (stockReel <= stockMini)
                    {
                        e.Appearance.ForeColor = Color.Red; 
                    }
                    else
                    {
                        e.Appearance.ForeColor = Color.Black;
                    }
                }
            }
        }
        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            frmSites frmsite = new frmSites(this,dbPrincipale);
            frmsite.ShowDialog();
            
        }
        private string GetArticleDescription(string arRef)
        {
            string description = string.Empty;
            string query = "SELECT AR_Design FROM F_ARTICLE WHERE AR_Ref = @AR_Ref"; // Remplace "Articles" par le nom exact de ta table

            using (SqlConnection conn = new SqlConnection(connectionString)) // Assure-toi que `connectionString` est bien défini
            {
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.Add("@AR_Ref", SqlDbType.NVarChar, 50).Value = arRef;

                    try
                    {
                        conn.Open();
                        object result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            description = result.ToString();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MethodBase m = MethodBase.GetCurrentMethod();
                        MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            return description;
        }
        private string GetFRNS(string arRef)
        {
            string frns = string.Empty;
            string query = "SELECT CT_Num FROM F_ARTFOURNISS WHERE AR_Ref = @AR_Ref"; // Remplace "Articles" par le nom exact de ta table

            using (SqlConnection conn = new SqlConnection(connectionString)) // Assure-toi que `connectionString` est bien défini
            {
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.Add("@AR_Ref", SqlDbType.NVarChar, 50).Value = arRef;

                    try
                    {
                        conn.Open();
                        object result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            frns = result.ToString();
                        }
                        else
                        {
                            frns = "";
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MethodBase m = MethodBase.GetCurrentMethod();
                        MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            return frns;
        }

        private void btnEditLigne_Click(object sender, EventArgs e)
        {
            string _currentDocPieceNo = dopiecetxt.Text;
            InitializeGrid(gcLigneEdit, _currentDocPieceNo);
        }

        private Guid getcbCreationUserGuid(string usermail)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(@"SELECT PROT_Guid  --REPLACE(PROT_Guid,'{','')  
                                                                FROM 
                                                                F_PROTECTIONCIAL 
                                                                WHERE PROT_EMail = @usermail", conn))
                {
                    cmd.Parameters.Add("@usermail", SqlDbType.NVarChar, 256).Value = usermail;
                    {
                        conn.Open();
                        Guid? result = cmd.ExecuteScalar() as Guid?;
                        if (result.HasValue)
                            return result.Value;
                    }
                }
            }
            return Guid.Empty;
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
        private void btnValider_Click(object sender, EventArgs e)
        {
            if (txtCours.Text != "" || txtCours.Text == "0,00")
            {
                decimal totalPoids = 0;

                if (gvLigneEdit.RowCount > 0)
                {
                    for (int i = 0; i < gvLigneEdit.RowCount; i++)
                    {
                        object value = gvLigneEdit.GetRowCellValue(i, "DL_PoidsNet");

                        if (value != null && value != DBNull.Value)
                            totalPoids += Convert.ToDecimal(value);
                    }

                    if (txt_poids.Text != "")
                    {
                        decimal totPoids = Convert.ToDecimal(txt_poids.Text.ToString());

                        if (totPoids != totalPoids)
                        {
                            MessageBox.Show("Le poids total FRET est différent au total des poids dans la ligne", "Erreur", MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                        }
                        else
                        {

                            try
                            {
                                if (!ValiderChampsObligatoires())
                                    return;

                                F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(dopiecetxt.Text, _listeDocs);
                                string _currentDocPieceNo = dopiecetxt.Text;
                                if (doc != null)
                                {
                                    if (tester_cloturer(dopiecetxt.Text))
                                    {
                                        MessageBox.Show(
                                            "Ce document est déjà clôturé, vous ne pouvez plus modifier son contenu!!!!",
                                            "Modification bloquée",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );
                                    }
                                    else
                                    {
                                        if (dopiecetxt.Text.ToString().StartsWith("AFA"))
                                        {
                                            bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                                            if (autorise)
                                            {
                                                UpdateFDOCENTETE();
                                                _ucDocuments.RafraichirDonnees();
                                                gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                                StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                                MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                //frmSites frmsite = new frmSites(this,dbPrincipale);
                                               // frmsite.ShowDialog();
                                            }
                                            else
                                            {
                                                MessageBox.Show(
                                                    "Vous n'avez pas l'autorisation de mettre à jour une facture !",
                                                    "Modification bloquée",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error
                                                );

                                            }
                                        }
                                        else if (dopiecetxt.Text.ToString().StartsWith("APA"))
                                        {
                                            bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                                            if (autorise)
                                            {
                                                if (lkStatut.Text != "Accepté")
                                                {
                                                    UpdateFDOCENTETE();
                                                    _ucDocuments.RafraichirDonnees();
                                                    gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                                    StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                                    MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                    frmSites frmsite = new frmSites(this,dbPrincipale);
                                                    frmsite.ShowDialog();
                                                    list_depot();
                                                    ExecuteStockAlert();
                                                    ChargerArtFrns();
                                                }
                                                else
                                                {
                                                    MessageBox.Show(
                                                    "Vous n'avez pas l'autorisation de modifier ce statut, seul DG a le droit !",
                                                    "Modification bloquée",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error
                                                );
                                                }
                                            }
                                            else
                                            {
                                                MessageBox.Show(
                                                    "Vous n'avez pas l'autorisation de mettre à jour un projet d'achat !",
                                                    "Modification bloquée",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error
                                                );
                                            }
                                        }
                                        else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
                                        {
                                            bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                                            if (autorise)
                                            {
                                                UpdateFDOCENTETE();
                                                _ucDocuments.RafraichirDonnees();
                                                gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                                StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                                MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                frmSites frmsite = new frmSites(this,dbPrincipale);
                                                frmsite.ShowDialog();
                                                list_depot();
                                                ExecuteStockAlert();
                                                ChargerArtFrns();
                                            }
                                            else
                                            {
                                                MessageBox.Show(
                                                    "Vous n'avez pas l'autorisation de mettre à jour un bon de commande !",
                                                    "Modification bloquée",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error
                                                );
                                            }
                                        }
                                        else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
                                        {
                                            bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                                            if (autorise)
                                            {
                                                UpdateFDOCENTETE();
                                                _ucDocuments.RafraichirDonnees();
                                                gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                                StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                                MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                frmSites frmsite = new frmSites(this,dbPrincipale);
                                                frmsite.ShowDialog();
                                                list_depot();
                                                ExecuteStockAlert();
                                                ChargerArtFrns();
                                            }
                                            else
                                            {
                                                MessageBox.Show(
                                                    "Vous n'avez pas l'autorisation de mettre à jour un Bon de réception !",
                                                    "Modification bloquée",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error
                                                );
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    var test = _context.F_DOCENTETE.FirstOrDefault(d => d.DO_Piece == _currentDocPieceNo);

                                    if (test == null)
                                    {
                                        InsertFDOCENTETE();
                                        _ucDocuments.RafraichirDonnees();
                                        gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                        StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                        MessageBox.Show("Insertion terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        frmSites frmsite = new frmSites(this,dbPrincipale);
                                        frmsite.ShowDialog();
                                        list_depot();
                                        ExecuteStockAlert();
                                        ChargerArtFrns();
                                    }
                                    else
                                    {
                                        UpdateFDOCENTETE();
                                        _ucDocuments.RafraichirDonnees();
                                        gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                        MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        frmSites frmsite = new frmSites(this,dbPrincipale);
                                        frmsite.ShowDialog();
                                        list_depot();
                                        ExecuteStockAlert();
                                        ChargerArtFrns();
                                    }
                                }
                            }

                            catch (System.Exception ex)
                            {
                                MethodBase m = MethodBase.GetCurrentMethod();
                                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                                            MessageBoxIcon.Error);
                            }
                        }
                    }
                    else
                    {
                        if (!ValiderChampsObligatoires())
                            return;

                        F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(dopiecetxt.Text, _listeDocs);
                        string _currentDocPieceNo = dopiecetxt.Text;
                        if (doc != null)
                        {
                            if (tester_cloturer(dopiecetxt.Text))
                            {
                                MessageBox.Show(
                                    "Ce document est déjà clôturé, vous ne pouvez plus modifier son contenu!!!!",
                                    "Modification bloquée",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                            }
                            else
                            {
                                if (dopiecetxt.Text.ToString().StartsWith("AFA"))
                                {
                                    bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                                    if (autorise)
                                    {
                                        UpdateFDOCENTETE();
                                        _ucDocuments.RafraichirDonnees();
                                        gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                        StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                        MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        frmSites frmsite = new frmSites(this,dbPrincipale);
                                        frmsite.ShowDialog();
                                        list_depot();
                                        ExecuteStockAlert();
                                        ChargerArtFrns();
                                    }
                                    else
                                    {
                                        MessageBox.Show(
                                            "Vous n'avez pas l'autorisation de mettre à jour une facture !",
                                            "Modification bloquée",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );

                                    }
                                }
                                else if (dopiecetxt.Text.ToString().StartsWith("APA"))
                                {
                                    bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                                    if (autorise)
                                    {
                                        if (lkStatut.Text != "Accepté")
                                        {
                                            UpdateFDOCENTETE();
                                            _ucDocuments.RafraichirDonnees();
                                            gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                            StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                            MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            frmSites frmsite = new frmSites(this,dbPrincipale);
                                            frmsite.ShowDialog();
                                            list_depot();
                                            ExecuteStockAlert();
                                            ChargerArtFrns();
                                        }
                                        else
                                        {
                                            MessageBox.Show(
                                            "Vous n'avez pas l'autorisation de modifier ce statut, seul DG a le droit !",
                                            "Modification bloquée",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show(
                                            "Vous n'avez pas l'autorisation de mettre à jour un projet d'achat !",
                                            "Modification bloquée",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );
                                    }
                                }
                                else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
                                {
                                    bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                                    if (autorise)
                                    {
                                        UpdateFDOCENTETE();
                                        _ucDocuments.RafraichirDonnees();
                                        gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                        StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                        MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        frmSites frmsite = new frmSites(this,dbPrincipale);
                                        frmsite.ShowDialog();
                                        list_depot();
                                        ExecuteStockAlert();
                                        ChargerArtFrns();
                                    }
                                    else
                                    {
                                        MessageBox.Show(
                                            "Vous n'avez pas l'autorisation de mettre à jour un bon de commande !",
                                            "Modification bloquée",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );
                                    }
                                }
                                else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
                                {
                                    bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                                    if (autorise)
                                    {
                                        UpdateFDOCENTETE();
                                        _ucDocuments.RafraichirDonnees();
                                        gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                        StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                        MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        frmSites frmsite = new frmSites(this,dbPrincipale);
                                        frmsite.ShowDialog();
                                        list_depot();
                                        ExecuteStockAlert();
                                        ChargerArtFrns();
                                    }
                                    else
                                    {
                                        MessageBox.Show(
                                            "Vous n'avez pas l'autorisation de mettre à jour un Bon de réception !",
                                            "Modification bloquée",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );
                                    }
                                }
                            }
                        }
                        else
                        {
                            var test = _context.F_DOCENTETE.FirstOrDefault(d => d.DO_Piece == _currentDocPieceNo);

                            if (test == null)
                            {
                                InsertFDOCENTETE();
                                _ucDocuments.RafraichirDonnees();
                                gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                MessageBox.Show("Insertion terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                frmSites frmsite = new frmSites(this,dbPrincipale);
                                frmsite.ShowDialog();
                                list_depot();
                                ExecuteStockAlert();
                                ChargerArtFrns();
                            }
                            else
                            {
                                UpdateFDOCENTETE();
                                _ucDocuments.RafraichirDonnees();
                                gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                frmSites frmsite = new frmSites(this,dbPrincipale);
                                frmsite.ShowDialog();
                                list_depot();
                                ExecuteStockAlert();
                                ChargerArtFrns();
                            }
                        }
                    }
                }
                else
                {
                    if (!ValiderChampsObligatoires())
                        return;

                    F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(dopiecetxt.Text, _listeDocs);
                    string _currentDocPieceNo = dopiecetxt.Text;
                    if (doc != null)
                    {
                        if (tester_cloturer(dopiecetxt.Text))
                        {
                            MessageBox.Show(
                                "Ce document est déjà clôturé, vous ne pouvez plus modifier son contenu!!!!",
                                "Modification bloquée",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                        else
                        {
                            if (dopiecetxt.Text.ToString().StartsWith("AFA"))
                            {
                                bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                                if (autorise)
                                {
                                    UpdateFDOCENTETE();
                                    _ucDocuments.RafraichirDonnees();
                                    gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                    StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                    MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    frmSites frmsite = new frmSites(this,dbPrincipale);
                                    frmsite.ShowDialog();
                                    list_depot();
                                    ExecuteStockAlert();
                                    ChargerArtFrns();
                                }
                                else
                                {
                                    MessageBox.Show(
                                        "Vous n'avez pas l'autorisation de mettre à jour une facture !",
                                        "Modification bloquée",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error
                                    );

                                }
                            }
                            else if (dopiecetxt.Text.ToString().StartsWith("APA"))
                            {
                                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                                if (autorise)
                                {
                                    if (lkStatut.Text != "Accepté")
                                    {
                                        UpdateFDOCENTETE();
                                        _ucDocuments.RafraichirDonnees();
                                        gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                        StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                        MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        frmSites frmsite = new frmSites(this,dbPrincipale);
                                        frmsite.ShowDialog();

                                        list_depot();
                                        ExecuteStockAlert();
                                        ChargerArtFrns();
                                    }
                                    else
                                    {
                                        MessageBox.Show(
                                        "Vous n'avez pas l'autorisation de modifier ce statut, seul DG a le droit !",
                                        "Modification bloquée",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error
                                    );
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(
                                        "Vous n'avez pas l'autorisation de mettre à jour un projet d'achat !",
                                        "Modification bloquée",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error
                                    );
                                }
                            }
                            else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
                            {
                                bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                                if (autorise)
                                {
                                    UpdateFDOCENTETE();
                                    _ucDocuments.RafraichirDonnees();
                                    gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                    StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                    MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    frmSites frmsite = new frmSites(this,dbPrincipale);
                                    frmsite.ShowDialog();
                                    list_depot();
                                    ExecuteStockAlert();
                                    ChargerArtFrns();
                                }
                                else
                                {
                                    MessageBox.Show(
                                        "Vous n'avez pas l'autorisation de mettre à jour un bon de commande !",
                                        "Modification bloquée",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error
                                    );
                                }
                            }
                            else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
                            {
                                bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                                if (autorise)
                                {
                                    UpdateFDOCENTETE();
                                    _ucDocuments.RafraichirDonnees();
                                    gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                                    StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                                    MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    frmSites frmsite = new frmSites(this,dbPrincipale);
                                    frmsite.ShowDialog();
                                    list_depot();
                                    ExecuteStockAlert();
                                    ChargerArtFrns();
                                }
                                else
                                {
                                    MessageBox.Show(
                                        "Vous n'avez pas l'autorisation de mettre à jour un Bon de réception !",
                                        "Modification bloquée",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error
                                    );
                                }
                            }
                        }
                    }
                    else
                    {
                        var test = _context.F_DOCENTETE.FirstOrDefault(d => d.DO_Piece == _currentDocPieceNo);

                        if (test == null)
                        {
                            InsertFDOCENTETE();
                            _ucDocuments.RafraichirDonnees();
                            gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                            StatutActuel = Convert.ToInt32(lkStatut.EditValue);
                            MessageBox.Show("Insertion terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            frmSites frmsite = new frmSites(this,dbPrincipale);
                            frmsite.ShowDialog();
                            list_depot();
                            ExecuteStockAlert();
                            ChargerArtFrns();
                        }
                        else
                        {
                            UpdateFDOCENTETE();
                            _ucDocuments.RafraichirDonnees();
                            gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                            MessageBox.Show("Modification terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            frmSites frmsite = new frmSites(this,dbPrincipale);
                            frmsite.ShowDialog();
                            list_depot();
                            ExecuteStockAlert();
                            ChargerArtFrns();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("La saisie du cours de devise est obligatoire!!!!", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadCodeTaxe()
        {

            string query = "SELECT DISTINCT TA_Intitule, TA_Code FROM F_TAXE"; // Remplacez par votre table et colonnes

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(query, conn);
                    DataTable dataTable = new DataTable();
                    dataAdapter.Fill(dataTable);

                    //lkCodeTaxe.Properties.DataSource = dataTable;
                    //lkCodeTaxe.Properties.DisplayMember = "TA_Intitule";
                    //lkCodeTaxe.Properties.ValueMember = "TA_Code";
                    if (dataTable.Rows.Count == 0)
                    {
                        // Créer une ligne par défaut
                        DataRow defaultRow = dataTable.NewRow();
                        defaultRow["TA_Code"] = 0;                 // valeur par défaut
                        defaultRow["TA_Intitule"] = "Aucune taxe"; // texte par défaut
                        dataTable.Rows.Add(defaultRow);
                    }

                    // Assigner la source de données
                    lkCodeTaxe.Properties.DataSource = dataTable;
                    lkCodeTaxe.Properties.DisplayMember = "TA_Intitule";
                    lkCodeTaxe.Properties.ValueMember = "TA_Code";

                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private bool ValiderChampsObligatoires()
        {

            if (lkEdFrns.EditValue == null)
            {
                MessageBox.Show("Veuillez sélectionner un fournisseur.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lkEdFrns.Focus();
                return false;
            }

            if (lkExpedition.EditValue == null)
            {
                MessageBox.Show("Veuillez sélectionner une expédition.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lkExpedition.Focus();
                return false;
            }
            if (lkEdCollaborateur.EditValue == null)
            {
                MessageBox.Show("L'acheteur est obligatoire.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lkEdCollaborateur.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(lkCodeTaxe.Text))
            {
                MessageBox.Show("Le code taxe est obligatoire.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lkCodeTaxe.Focus();
                return false;
            }
            //lkDepot
            if (lkDepot.Text == null || lkDepot.Text == "")
            {
                MessageBox.Show("Définir le dépôt.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lkDepot.Focus();
                return false;
            }
            
            return true;
        }

        public int deviseID;
        public decimal ct_taux2;
        string CodeTaxe = string.Empty;
        string _prefix = string.Empty;
        private void UpdateFDOCENTETE()
        {
            try
            {
                Guid cbCreationUser = getcbCreationUserGuid(FrmConnex.mailuser);
                string caNum = comboBoxAffaire.Text;
                int caisseNumber = 0;
                int caissierNumber = 0;
                string expeditIntitule = lkExpedition.Text;


                DateTime dateLivrPrevu;
                DateTime minSqlDate = new DateTime(1753, 1, 1);

                if (datelivrprev.EditValue == null)
                {
                    dateLivrPrevu = minSqlDate;
                }
                else
                {
                    DateTime parsedDate = Convert.ToDateTime(datelivrprev.EditValue);

                    if (parsedDate < minSqlDate)
                        dateLivrPrevu = minSqlDate;
                    else
                        dateLivrPrevu = parsedDate;
                }


                DateTime dateLivrReal;

                if (datelivrprev.EditValue == null ||
                    !(datelivrprev.EditValue is DateTime) ||
                    (DateTime)datelivrprev.EditValue < minSqlDate)
                {
                    dateLivrReal = minSqlDate;
                }
                else
                {
                    dateLivrReal = (DateTime)datelivrprev.EditValue;
                }


                string reference = txtDoRef.Text;
                string dO_Coord01 = txtCoord1.Text;
                
                //string representant = _listeAcheteurs[lkEdCollaborateur.ItemIndex].ToString();
                int CO_NO = (int)cono;
                lkEdCollaborateur.EditValue = CO_NO;
                F_COLLABORATEUR representant = _listeCollaborateurs[lkEdCollaborateur.ItemIndex];
                F_COMPTET frns = _listeFrns[lkEdFrns.ItemIndex];
                string ctnum = frns.CT_Num;
                int numeroDepot = Convert.ToInt32(lkDepot.EditValue);
                string _typeDocument = ucDocuments.a_type;
                string _currentDocPieceNo = dopiecetxt.Text;
                int numExpedition = doexpedit;
                decimal docours = decimal.Parse(txtCours.Text);

                int number = int.Parse(dopiecetxt.Text.Substring(dopiecetxt.Text.Length - 4));
                decimal DoTaxe1 = string.IsNullOrEmpty(doTaxe1txt.Text)? 0: Convert.ToDecimal(doTaxe1txt.Text);

                short DoStatut = Convert.ToInt16(lkStatut.EditValue);

                if (fDocenteteToModif == null)
                {
                    _f_DOCENTETEService.UpdateProprietesF_DOCENTETE(_currentDocPieceNo,
    dateLivrPrevu, dateLivrReal, reference, caNum, CO_NO,
    (short?)numExpedition, expeditIntitule, dO_Coord01, numeroDepot, DoTaxe1, DoStatut, ctnum, docours);
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        string ExpeditionID = string.Empty;
        private void InsertFDOCENTETE()
        {
            Guid cbCreationUser = getcbCreationUserGuid(FrmConnex.mailuser);
                string caNum = comboBoxAffaire.Text;
                int caisseNumber = 0;
                int caissierNumber = 0;
                string expeditIntitule = ExpeditionID;
                short? numExpedition = Convert.ToInt16(lkExpedition.EditValue); // Fix: Convert 'int' to 'short?'

                DateTime dateLivrPrevu;
                if (datelivrprev.EditValue == null)
                    dateLivrPrevu = new DateTime(1753, 01, 01);
                else
                    dateLivrPrevu = (DateTime)datelivrprev.EditValue;
                DateTime dateLivrReal;
                if (datelivrprev.EditValue == null)
                    dateLivrReal = new DateTime(1753, 01, 01);
                else
                    dateLivrReal = (DateTime)datelivrprev.EditValue;
                string reference = txtDoRef.Text;
                string dO_Coord01 = txtCoord1.Text;

                F_COLLABORATEUR representant = _listeCollaborateurs[lkEdCollaborateur.ItemIndex];

                F_COMPTET frns = _listeFrns[lkEdFrns.ItemIndex];
                string ctnum = frns.CT_Num;
                int numeroDepot = Convert.ToInt32(lkDepot.EditValue);
                string _typeDocument = "Projet d'achat";
                string _currentDocPieceNo = dopiecetxt.Text;
                int number = int.Parse(dopiecetxt.Text.Substring(dopiecetxt.Text.Length - 4));
                decimal DoTaxe1 = 0m;
                decimal.TryParse(doTaxe1txt.Text, out DoTaxe1);
                decimal doCours = decimal.Parse(txtCours.Text);

                if (fDocenteteToModif == null)
                {
                    _f_DOCENTETEService.InsertNewF_DOCENTETE(_typeDocument, _currentDocPieceNo, frns,
                        numExpedition, caNum, caisseNumber,
                        caissierNumber, expeditIntitule, dateLivrPrevu,
                        dateLivrReal, reference, dO_Coord01,
                        (int)representant.CO_No, numeroDepot, deviseID,
                        ct_taux2, CodeTaxe, _prefix,
                        number, DoTaxe1, doCours);
                }
        }
        public void TransformFDOCENTETE(string doctype)
        {
            try
            {

                Guid cbCreationUser = FrmMdiParent._id_user;   //getcbCreationUserGuid(FrmConnex.mailuser);
                string caNum = comboBoxAffaire.Text;
                int caisseNumber = 0;
                int caissierNumber = 0;
                short newStatut = 0;
                string expeditIntitule = lkExpedition.Text;
                _listeDocs = _context.F_DOCENTETE.ToList();
                F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(ucDocuments.doPiece, _listeDocs);
                int cbMarqSource = doc.cbMarq;
                DateTime dateLivrPrevu;
                DateTime minSqlDate = new DateTime(1753, 1, 1);

                if (datelivrprev.EditValue == null)
                {
                    dateLivrPrevu = minSqlDate;
                }
                else
                {
                    DateTime parsedDate = Convert.ToDateTime(datelivrprev.EditValue);

                    if (parsedDate < minSqlDate)
                        dateLivrPrevu = minSqlDate;
                    else
                        dateLivrPrevu = parsedDate;
                }


                DateTime dateLivrReal;

                if (datelivrprev.EditValue == null ||
                    !(datelivrprev.EditValue is DateTime) ||
                    (DateTime)datelivrprev.EditValue < minSqlDate)
                {
                    dateLivrReal = minSqlDate;
                }
                else
                {
                    dateLivrReal = (DateTime)datelivrprev.EditValue;
                }


                string reference = txtDoRef.Text;
                string dO_Coord01 = txtCoord1.Text;

                string representant = _listeAcheteurs[lkEdCollaborateur.ItemIndex].ToString();
                int CO_NO = (int)cono;

                F_COMPTET frns = _listeFrns[lkEdFrns.ItemIndex];
                string ctnum = frns.CT_Num;
                int numeroDepot = Convert.ToInt32(lkDepot.EditValue);
                string _typeDocument = ucDocuments.a_type;
                string _currentDocPieceNo = dopiecetxt.Text;
                string newDocPieceNo = TransformNoPiece(doctype, dopiecetxt.Text);
                int numExpedition = doexpedit;

                int number = int.Parse(dopiecetxt.Text.Substring(dopiecetxt.Text.Length - 4));
                decimal DoTaxe1 = Convert.ToDecimal(doTaxe1txt.Text);
                //short DoStatut = Convert.ToInt16(lkStatut.EditValue);
                short newDoType = Convert.ToInt16(TransformDoType(doctype, dopiecetxt.Text));
                DateTime newDoDate = DateTime.Now;

                if (fDocenteteToModif == null)
                {
                    _f_DOCENTETEService.TransformF_DOCENTETE(cbMarqSource, newDoType, _currentDocPieceNo, newDocPieceNo,

    dateLivrPrevu, dateLivrReal, reference, caNum, CO_NO,
    (short?)numExpedition, expeditIntitule, dO_Coord01, numeroDepot, DoTaxe1, newStatut, newDoDate);
                }


                ///////////////////////////////////F_DOCLIGNE

                using (_context)
                {
                    List<F_DOCLIGNE> listeDoclignesToUpdate = _f_DOCLIGNERepository.GetAll_F_DOCLIGNE_Of_DOCENTETE(_currentDocPieceNo);


                    if (listeDoclignesToUpdate != null)
                    {
                        foreach (var ligne in listeDoclignesToUpdate)
                        {
                            bool retenu = (bool)ligne.Retenu;
                            if (retenu == true) // ou simplement: if (ligne.Retenu)
                            {
                                int dl_No = (int)ligne.DL_No;
                                _f_DOCLIGNEService.TransformF_DOCLIGNE(newDoType, newDocPieceNo, dl_No);
                            }
                        }
                    }
                }
                //////////////////////////////////////////////////////////////////////////////////////////////////
                string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";

                string UPDATESql = @"UPDATE F_MODE_PAIEMENT SET do_piece=@do_pieces WHERE do_piece=@do_piece";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                using (SqlCommand cmd = new SqlCommand(UPDATESql, conn))
                {
                    cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                        .Value = dopiecetxt.Text.Trim();

                    cmd.Parameters.Add("@do_pieces", SqlDbType.VarChar, 20)
                        .Value = newDocPieceNo;

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }

                string UPDATESql1 = @"UPDATE F_PACKING_LIST SET dopiece=@do_pieces WHERE dopiece=@do_piece";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                using (SqlCommand cmd = new SqlCommand(UPDATESql1, conn))
                {
                    cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                        .Value = dopiecetxt.Text.Trim();

                    cmd.Parameters.Add("@do_pieces", SqlDbType.VarChar, 20)
                        .Value = newDocPieceNo;

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }

                this.Close();
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string TransformNoPiece(string doctype, string docpiece)
        {
            string newdocpiece = "";
            switch (doctype)
            {
                case "Facture":
                    newdocpiece = dopiece.Replace(docpiece.Substring(0, 3), "AFA");
                    break;

                case "Bon de livraison":
                    newdocpiece = dopiece.Replace(docpiece.Substring(0, 3), "ABL");
                    break;

                case "Bon de commande":
                    newdocpiece = dopiece.Replace(docpiece.Substring(0, 3), "ABC");
                    break;

                case "Bon de réception":
                    newdocpiece = dopiece.Replace(docpiece.Substring(0, 3), "ABR");
                    break;
            }

            var f_fs = _context.F_FRETS.FirstOrDefault(x => x.DO_PIECE == dopiece);

            if (f_fs != null)
            {
                f_fs.DO_PIECE =newdocpiece;
                _context.SaveChanges();
            }

            return newdocpiece;
        }
        private string TransformDoType(string doctype, string docpiece)
        {
            int newDoType = 10;
            switch (doctype)
            {
                case "Préparation de commande":
                    newDoType = 11;
                    break;
                case "Bon de commande":
                    newDoType = 12;
                    break;
                case "Bon de livraison":
                    newDoType = 13;
                    break;
                case "Bon de retour":
                    newDoType = 14;
                    break;
                case "Facture":
                    newDoType = 16;
                    break;
                case "Bon de réception":
                    newDoType = 18;
                    break;
                default:
                    throw new ArgumentException("Invalid document type", nameof(doctype));
            }
            return newDoType.ToString(); // Convert int to string to fix CS0029 error
        }
        private void lkCodeTaxe_EditValueChanged(object sender, EventArgs e)
        {
            CodeTaxe = lkCodeTaxe.EditValue.ToString();
            if (CodeTaxe == "0")
            {
                doTaxe1txt.Text = string.Empty;
                return;
            }

            string codeTaxe = lkCodeTaxe.EditValue.ToString();

            F_TAXE taxe = _f_TAXERepository?.Get_F_TAXE_By_TA_Code(codeTaxe);
            decimal taux = taxe?.TA_Taux ?? 0;
            doTaxe1txt.Text = taux.ToString("0.##");




        }

        private void hyperlinkLabelControl4_Click(object sender, EventArgs e)
        {
            gvLigneEdit.Columns["DL_Frais"].Visible = false;
            gvLigneEdit.Columns["DL_Taxe1"].Visible = false;
            gvLigneEdit.Columns["AF_RefFourniss"].Visible = false;
            gvLigneEdit.Columns["DL_Remise01REM_Valeur"].Visible = false;
            gvLigneEdit.Columns["DL_Frais"].Visible = false;
            gvLigneEdit.Columns["AR_RefCompose"].Visible = false;
            gvLigneEdit.Columns["DL_NonLivre"].Visible = false;
            gvLigneEdit.Columns["DL_PUDevise"].Visible = false;
            gvLigneEdit.Columns["DL_MontantHT"].Visible = false;
            gvLigneEdit.Columns["DL_MontantTTC"].Visible = false;
            gvLigneEdit.Columns["Retenu"].Visible = false;
            gvLigneEdit.Columns["Action"].Visible = false;
            gvLigneEdit.Columns["Validation"].Visible = false;
            gcLigneEdit.ShowPrintPreview();
            gvLigneEdit.Columns["DL_Frais"].Visible = true;
            gvLigneEdit.Columns["DL_Taxe1"].Visible = true;
            gvLigneEdit.Columns["AF_RefFourniss"].Visible = true;
            gvLigneEdit.Columns["DL_Remise01REM_Valeur"].Visible = true;
            gvLigneEdit.Columns["AR_RefCompose"].Visible = true;
            gvLigneEdit.Columns["DL_NonLivre"].Visible = true;
            gvLigneEdit.Columns["DL_PUDevise"].Visible = true;
            gvLigneEdit.Columns["DL_MontantHT"].Visible = true;
            gvLigneEdit.Columns["DL_MontantTTC"].Visible = true;
            gvLigneEdit.Columns["Retenu"].Visible = true;
            gvLigneEdit.Columns["Action"].Visible = true;
            gvLigneEdit.Columns["Validation"].Visible = true;


        }

        private void dopiecetxt_EditValueChanged(object sender, EventArgs e)
        {
            //ChargerLignes(dopiecetxt.Text);
            Lignes.AfficherLignes(gcLigneEdit, dopiecetxt.Text);



            Lignes.cacherColonnes(gvLigneEdit);
            SetupGridView();
            dopiece = dopiecetxt.Text;
            if(dopiece != "")
            {
                LoadDocLie(dopiece);
            }
            
                
        }

        private void Editor_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit editor = sender as CheckEdit;
            int rowHandle = gvLigneEdit.FocusedRowHandle;

            if (rowHandle >= 0)
            {
                gvLigneEdit.SetRowCellValue(rowHandle, "Retenu", editor.Checked);
                gvLigneEdit.UpdateSummary(); // ⬅️ force le recalcul du footer
                UPdateLigne(rowHandle);
            }
        }

        private bool isUpdating = false;

        private void gvLigneEdit_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (isUpdating) return; // éviter boucle infinie
            isUpdating = true;

            GridView view = sender as GridView;

            if (e.Column.FieldName == "PU par tonne")
            {
                if (decimal.TryParse(view.GetRowCellValue(e.RowHandle, "PU par tonne")?.ToString(), out decimal prixTonne))
                {
                    decimal val = prixTonne / 1000;
                    view.SetRowCellValue(e.RowHandle, "DL_PrixUnitaire", val); // prix par kg
                }
            }
            else if (e.Column.FieldName == "DL_PrixUnitaire")
            {
                if (decimal.TryParse(view.GetRowCellValue(e.RowHandle, "DL_PrixUnitaire")?.ToString(), out decimal prixKg))
                {
                    view.SetRowCellValue(e.RowHandle, "PU par tonne", prixKg * 1000); // prix par tonne
                }
            }

            isUpdating = false;

            gvLigneEdit.PostEditor();

            if (e.Column.FieldName == "AR_Ref" && e.Value != null && !string.IsNullOrWhiteSpace(e.Value.ToString()))
            {
                string selectedAR_Ref = e.Value.ToString();

                if (e.RowHandle != DevExpress.XtraGrid.GridControl.NewItemRowHandle)
                {
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DL_Design", GetArticleDescription(selectedAR_Ref));
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Action", "Remove");
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Retenu", 1);
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Validation", "Update");
                    //gvLigneEdit.SetRowCellValue(e.RowHandle, "Insertion", "Add");
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DL_PrixUnitaire", GetArticlePU(selectedAR_Ref));
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "CT_Num", GetFRNS(selectedAR_Ref));

                    gvLigneEdit.BestFitColumns();
                }
                else
                {
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DL_Design", GetArticleDescription(selectedAR_Ref));
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Action", "Remove");
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Retenu", 1);
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Validation", "Update");
                    //gvLigneEdit.SetRowCellValue(e.RowHandle, "Insertion", "Add");
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DL_PrixUnitaire", GetArticlePU(selectedAR_Ref));
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "CT_Num", GetFRNS(selectedAR_Ref));

                    gvLigneEdit.BestFitColumns();
                }
            }
            if (e.Column.FieldName == "DL_Qte" || e.Column.FieldName == "DL_Remise01REM_Valeur" || e.Column.FieldName == "DL_PrixUnitaire"
                || e.Column.FieldName == "DL_Taxe1" || e.Column.FieldName == "DL_PUDevise" || e.Column.FieldName == "DL_PoidsNet" || e.Column.FieldName == "DL_Frais")
            {
                //GridView view = sender as GridView;

                // Récupérer les valeurs
                object prixUnitaireObj = view.GetRowCellValue(e.RowHandle, "DL_PrixUnitaire");
                object quantiteObj = view.GetRowCellValue(e.RowHandle, "DL_Qte");
                int qte = 0;
                if (quantiteObj != null && quantiteObj != DBNull.Value && quantiteObj.ToString() != "")
                {
                    qte = Convert.ToInt32(quantiteObj);
                }
                else
                {
                    qte = 0;
                }

                object dlTaxe1obj = gvLigneEdit.GetRowCellValue(e.RowHandle, "DL_Taxe1");
                int dlTaxe1 = 0;
                if (dlTaxe1obj != null && dlTaxe1obj != DBNull.Value && dlTaxe1obj.ToString() != "")
                {
                    dlTaxe1 = Convert.ToInt32(dlTaxe1obj);
                }
                else
                {
                    dlTaxe1 = ucDocuments.doTaxe1;
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DO_Taxe1", dlTaxe1);
                }

                object dlPuDeviseobj = gvLigneEdit.GetRowCellValue(e.RowHandle, "DL_PuDevise");
                decimal dlPuDevise = 0;
                if (dlPuDeviseobj != null && dlPuDeviseobj != DBNull.Value && dlPuDeviseobj.ToString() != "")
                {
                    dlPuDevise = Convert.ToDecimal(dlPuDeviseobj);
                }
                else
                {
                    dlPuDevise = 0;
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DL_PuDevise", dlPuDevise);
                }

                object remiseObj = view.GetRowCellValue(e.RowHandle, "DL_Remise01REM_Valeur");
                object tauxTVAObj = view.GetRowCellValue(e.RowHandle, "DL_Taxe1");

                if (prixUnitaireObj != null && quantiteObj != null && remiseObj != null && tauxTVAObj != null &&
                    prixUnitaireObj != DBNull.Value && quantiteObj != DBNull.Value &&
                    remiseObj != DBNull.Value && tauxTVAObj != DBNull.Value)
                {
                    decimal prixUnitaire = Convert.ToDecimal(prixUnitaireObj);
                    int quantite = Convert.ToInt32(qte);
                    decimal remise = Convert.ToDecimal(remiseObj);
                    decimal tauxTVA = Convert.ToDecimal(tauxTVAObj);
                    decimal cours = Convert.ToDecimal(txtCours.Text);

                    // Appliquer la remise au prix unitaire
                    decimal prixNet = prixUnitaire * (1 - remise / 100);
                    decimal montantHT = prixNet * quantite * cours;
                    decimal montantTTC = montantHT * (1 + tauxTVA / 100);

                    view.SetRowCellValue(e.RowHandle, "DL_MontantHT", montantHT);
                    view.SetRowCellValue(e.RowHandle, "DL_MontantTTC", montantTTC);

                    gvLigneEdit.CustomUnboundColumnData += (sender, e) =>
                    {
                        if (e.IsGetData)
                        {
                            if (e.Column.FieldName == "FRET")
                            {
                                try
                                {
                                    decimal poidsNet = Convert.ToDecimal(gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "DL_PoidsNet"));
                                    decimal prix = Convert.ToDecimal(txt_prix.Text);
                                    decimal poids = Convert.ToDecimal(txt_poids.Text);

                                    decimal fret = (poidsNet * prix) / poids;
                                    if (fret < 0) fret = 0m;

                                    e.Value = fret.ToString("N2");
                                }catch(Exception er) { }
                            }
                            else if (e.Column.FieldName == "Total Frais")
                            {
                                try { 
                                    decimal fret = Convert.ToDecimal(gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "FRET"));
                                    decimal dlFrais = Convert.ToDecimal(gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "DL_Frais"));

                                    decimal tot_frais = dlFrais + fret;
                                    e.Value = tot_frais.ToString("N2");
                                }
                                catch (Exception er) { }
                            }
                        }
                    };
                }
            }
            
            gvLigneEdit.UpdateSummary();
            gvLigneEdit.RefreshData();
        }

        private decimal GetArticlePU(string arRef)
        {
            decimal puAch = 0;
            string query = @$"SELECT 
                              ISNULL(dbo.F_ARTFOURNISS.AF_PrixAch, 0) AS AF_PrixAch
                            FROM
                              dbo.F_ARTICLE
                              INNER JOIN dbo.F_ARTFOURNISS ON (dbo.F_ARTICLE.AR_Ref = dbo.F_ARTFOURNISS.AR_Ref)
                            WHERE dbo.F_ARTICLE.AR_Ref = @AR_Ref";

            using (SqlConnection conn = new SqlConnection(connectionString)) // Assure-toi que `connectionString` est bien défini
            {
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.Add("@AR_Ref", SqlDbType.NVarChar, 50).Value = arRef;

                    try
                    {
                        conn.Open();
                        object result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            puAch = (decimal)result;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MethodBase m = MethodBase.GetCurrentMethod();
                        MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
            return puAch;
        }
        private void gvLigneEdit_CellValueChanging(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {

            if (e.Column.FieldName == "AR_Ref" && e.Value != null && !string.IsNullOrWhiteSpace(e.Value.ToString()))
            {
                string selectedAR_Ref = e.Value.ToString();

                if (e.RowHandle != DevExpress.XtraGrid.GridControl.NewItemRowHandle)
                {
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DL_Design", GetArticleDescription(selectedAR_Ref));
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Action", "Remove");
                    //gvLigneEdit.SetRowCellValue(e.RowHandle, "Insertion", "Add");
                    //newRow["CT_Num"] = CT_Num;
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "CT_Num", GetFRNS(selectedAR_Ref));
                }
                else
                {
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "DL_Design", GetArticleDescription(selectedAR_Ref));
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "Action", "Remove");
                    //gvLigneEdit.SetRowCellValue(e.RowHandle, "Insertion", "Add");
                    gvLigneEdit.SetRowCellValue(e.RowHandle, "CT_Num", GetFRNS(selectedAR_Ref));
                }
            }

        }

        private void gvLigneEdit_ShowingEditor(object sender, CancelEventArgs e)
        {
            GridView view = sender as GridView;
            int row = view.FocusedRowHandle;

            // Seulement pour la colonne DL_PrixUnitaire
            if (view.FocusedColumn.FieldName == "DL_PrixUnitaire")
            {
                //e.Cancel = true; // ❌ Bloque l’éditeur DevExpress
                decimal prixActuel = Convert.ToDecimal(
                    view.GetRowCellValue(row, "DL_PrixUnitaire")
                );

                /*frmPrix f = new frmPrix(prixActuel);

                if (f.ShowDialog() == DialogResult.OK)
                {
                    // Écrire la valeur dans la cellule
                    view.SetRowCellValue(row, "DL_PrixUnitaire", f.Prix);
                    view.UpdateCurrentRow();
                }*/
            }
        }

        public int deno;
        private void lkDepot_EditValueChanged(object sender, EventArgs e)
        {
            if (lkDepot.EditValue != null && lkDepot.EditValue != DBNull.Value)
            {
                deno = Convert.ToInt32(lkDepot.EditValue);
            }
        }
        public int doexpedit;
        private void lkExpedition_EditValueChanged(object sender, EventArgs e)
        {
            if (lkExpedition.EditValue != null && lkExpedition.EditValue != DBNull.Value)
            {
                doexpedit = Convert.ToInt32(lkExpedition.EditValue);
            }
        }
        public int cono;
        private void lkEdCollaborateur_EditValueChanged(object sender, EventArgs e)
        {
            if (lkEdCollaborateur.EditValue != null && lkEdCollaborateur.EditValue != DBNull.Value)
            {
                cono = Convert.ToInt32(lkEdCollaborateur.EditValue);
            }

        }

        private void lister_packing(String cond)
        {
            lblval1.Text = "";
            string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";

            string query = "SELECT * FROM F_PACKING_LIST WHERE dopiece = @do_piece";

            using (SqlConnection conn = new SqlConnection(connectionString2))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                       .Value = dopiecetxt.Text.Trim();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows == true)
                            {
                                lblval1.Text = "1";
                                while (reader.Read())
                                {
                                    cmb_type.Text = reader.GetString(1).ToString();
                                    txtnbr.Text = reader.GetString(2);
                                    if (!reader.IsDBNull(3))
                                        dteta.Value = Convert.ToDateTime(reader.GetValue(3));

                                    if (!reader.IsDBNull(4))
                                        dtetd.Value = Convert.ToDateTime(reader.GetValue(4));
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void list_depot()
        {
            string connectionString3 =
                                   $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                   $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                   $"Connection Timeout=240;";

            string query_depot = "SELECT val_depot FROM DEPOT";

            using (SqlConnection conn = new SqlConnection(connectionString3))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query_depot, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows == true)
                            {
                                while (reader.Read())
                                {
                                    labelControl8.Text = reader.GetString(0).ToString();
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static int intcollaborateur;
        private Dictionary<int, decimal> puTonneValues = new Dictionary<int, decimal>();

        private void frmEditDocument_Load_1(object sender, EventArgs e)
        {
            lister_packing(dopiecetxt.Text);
            if (dopiecetxt.Text.StartsWith("AFA"))
            {
                groupControl4.Enabled = true;
            }
            else
            {
                groupControl4.Enabled = false;
            }

            lbl_val.Text = "";

            string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";


            list_depot();

            string query = "SELECT mode_paiement,date_echeance FROM F_MODE_PAIEMENT WHERE do_piece = @do_piece";

            using (SqlConnection conn = new SqlConnection(connectionString2))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                       .Value = dopiecetxt.Text.Trim();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if(reader.HasRows==true)
                            {
                                lbl_val.Text = "1";
                                while (reader.Read())
                                {
                                    cmbmdp.Text= reader.GetString(0).ToString();
                                    dtecheance.Value = reader.GetDateTime(1);
                                }
                            } 
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if (dopiecetxt.Text.Contains("ABC") || dopiecetxt.Text.Contains("AFA"))
            {
                groupControl3.Enabled = true;
            }
            else
            {
                groupControl3.Enabled = false;
            }

            treeList1.OptionsFind.ExpandNodesOnIncrementalSearch = true;

            var test_exist = _context.F_DOCENTETE.FirstOrDefault(x => x.DO_Piece == dopiecetxt.Text.ToString());

            if (test_exist != null)
            {
                ExecuteStockAlert();
                ChargerArtFrns();
                var list = _context.F_FRETS.FirstOrDefault(p => p.DO_PIECE == dopiecetxt.Text);

                if (list != null)
                {
                    try
                    {
                        if (list.DO_MONTANT.ToString() != "")
                        {
                            txt_prix.EditValue = list.DO_MONTANT.ToString("N2");
                            txt_poids.EditValue = list.DO_POIDS.ToString("N2");
                        }
                        else
                        {
                            txt_prix.EditValue = txt_poids.EditValue = "0";
                        }

                    }
                    catch (Exception ef) { }
                }
                else
                {
                    txt_prix.EditValue = txt_poids.EditValue = "0";
                }

                LkDevise_EditValueChanged(sender, e);

                //=== PU en tonne ===
                GridColumn pu_tonne = gvLigneEdit.Columns["PU par tonne"];
                if (pu_tonne == null)
                {
                    pu_tonne = gvLigneEdit.Columns.AddField("PU par tonne");
                    pu_tonne.Caption = "PU par tonne";
                    pu_tonne.UnboundType = DevExpress.Data.UnboundColumnType.Decimal;
                    pu_tonne.Visible = true;
                    pu_tonne.FieldName = "PU par tonne";
                    pu_tonne.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                }

                int indexDesign = gvLigneEdit.Columns["DL_Design"].VisibleIndex;
                pu_tonne.VisibleIndex = indexDesign + 1;

                RepositoryItemSpinEdit spin = new RepositoryItemSpinEdit();
                spin.IsFloatValue = true;
                spin.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
                spin.Mask.EditMask = "n2";
                spin.Mask.UseMaskAsDisplayFormat = true;

                // IMPORTANT : Gérer l'événement EditValueChanged du SpinEdit
                spin.EditValueChanged += (s, e) =>
                {
                    gvLigneEdit.PostEditor();  // Force la validation immédiate
                    gvLigneEdit.UpdateCurrentRow();
                };

                gvLigneEdit.GridControl.RepositoryItems.Add(spin);
                pu_tonne.ColumnEdit = spin;

                // Événement CustomUnboundColumnData
                gvLigneEdit.CustomUnboundColumnData += (sender, e) =>
                {
                    if (e.Column.FieldName == "PU par tonne")
                    {
                        int row = e.ListSourceRowIndex;

                        if (e.IsGetData)
                        {
                            if (puTonneValues.ContainsKey(row))
                                e.Value = puTonneValues[row];
                            else
                                e.Value = 0m;
                        }
                        else if (e.IsSetData)
                        {
                            if (e.Value != null && e.Value != DBNull.Value)
                            {
                                puTonneValues[row] = Convert.ToDecimal(e.Value);
                            }
                        }
                    }
                };

                // Forcer la mise à jour quand on change de cellule
                gvLigneEdit.FocusedColumnChanged += (sender, e) =>
                {
                    gvLigneEdit.PostEditor();
                    gvLigneEdit.UpdateCurrentRow();
                };

                // === UNITE ===
                GridColumn col_unite = gvLigneEdit.Columns["Unite"];
                if (col_unite == null)
                {
                    col_unite = gvLigneEdit.Columns.AddField("Unite");
                    col_unite.Caption = "Unité";
                    col_unite.UnboundType = DevExpress.Data.UnboundColumnType.Decimal;
                    col_unite.Visible = true;
                    col_unite.OptionsColumn.AllowEdit = false;
                    col_unite.OptionsColumn.ReadOnly = true;
                    col_unite.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                }

                // === FRET ===
                GridColumn col = gvLigneEdit.Columns["FRET"];
                if (col == null)
                {
                    col = gvLigneEdit.Columns.AddField("FRET");
                    col.Caption = "FRET";
                    col.UnboundType = DevExpress.Data.UnboundColumnType.Decimal;
                    col.Visible = true;
                    col.OptionsColumn.AllowEdit = true;
                    col.OptionsColumn.ReadOnly = false;
                }

                // === Total Frais ===
                GridColumn col1 = gvLigneEdit.Columns["TotalFrais"];

                if (col1 == null)
                {
                    col1 = gvLigneEdit.Columns.AddField("TotalFrais"); // 👈 SANS espace
                    col1.Caption = "Total Frais";                      // 👈 affiché
                    col1.UnboundType = DevExpress.Data.UnboundColumnType.Decimal;
                    col1.Visible = true;

                    col1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                    col1.DisplayFormat.FormatString = "N2";
                }

                GridColumn col_htdevise = gvLigneEdit.Columns["Montant Total en devise"];
                if (col_htdevise == null)
                {
                    col_htdevise = gvLigneEdit.Columns.AddField("Montant Total en devise");
                    col_htdevise.Caption = "Montant Total en devise";
                    col_htdevise.UnboundType = DevExpress.Data.UnboundColumnType.Decimal;
                    col_htdevise.Visible = true;
                    col_htdevise.BestFit();
                    col_htdevise.OptionsColumn.AllowEdit = false;
                    col_htdevise.OptionsColumn.ReadOnly = true;
                    col_htdevise.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                }

                // === POSITIONNEMENT ===
                int indexFrais = gvLigneEdit.Columns["DL_Frais"].VisibleIndex;
                col.VisibleIndex = indexFrais;
                col1.VisibleIndex = indexFrais + 2;

                int indexPrixUnitaire = gvLigneEdit.Columns["DL_PrixUnitaire"].VisibleIndex;
                col_unite.VisibleIndex = indexPrixUnitaire - 1;

                int indexMontantHT = gvLigneEdit.Columns["DL_MontantHT"].VisibleIndex;
                col_htdevise.VisibleIndex = indexMontantHT + 1;

                gvLigneEdit.CustomUnboundColumnData += (sender, e) =>
                {
                    if (e.IsGetData)
                    {
                        if (e.Column.FieldName == "FRET")
                        {
                            try
                            {
                                decimal poidsNet = Convert.ToDecimal(gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "DL_PoidsNet"));
                                if (txt_prix.Text != "" && Convert.ToDecimal(txt_prix.Text.ToString()) > 0)
                                {
                                    decimal prix = Convert.ToDecimal(txt_prix.Text);
                                    decimal poids = Convert.ToDecimal(txt_poids.Text);

                                    decimal fret = (poidsNet * prix) / poids;
                                    if (fret < 0) fret = 0m;

                                    e.Value = fret.ToString("N2");
                                }
                            }
                            catch (Exception es) { }
                        }
                        else if (e.Column.FieldName == "Montant Total en devise")
                        {
                            try
                            {
                                decimal PUHT = Convert.ToDecimal(gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "DL_MontantHT"));
                                if (txtCours.Text != "" && Convert.ToDecimal(txtCours.Text.ToString()) > 0)
                                {
                                    decimal cours_devise = Convert.ToDecimal(txtCours.Text);
                                    decimal m_devise = PUHT / cours_devise;
                                    if (m_devise < 0) m_devise = 0m;

                                    e.Value = m_devise.ToString("N2");

                                    col_htdevise.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
                                    col_htdevise.SummaryItem.DisplayFormat = "{0:n2}";
                                    gvLigneEdit.Appearance.FooterPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                                }
                            }
                            catch (Exception es) { }
                        }
                        else if (e.Column.FieldName == "TotalFrais" && e.IsGetData)
                        {
                            decimal fret = Convert.ToDecimal(gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "FRET"));
                            decimal dlFrais = 0m;
                            object val = gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "DL_Frais");
                            if (val != null && val != DBNull.Value)
                                dlFrais = Convert.ToDecimal(val);

                            decimal tot_frais = dlFrais + fret;
                            e.Value = tot_frais; // reste un decimal

                            //Récupération unité

                            try
                            {
                                string reference = gvLigneEdit.GetListSourceRowCellValue(e.ListSourceRowIndex, "AR_Ref").ToString();
                                var recup = _context.F_ARTFOURNISS
                                .FirstOrDefault(x => x.AR_Ref == reference);

                                if (recup != null)
                                {
                                    string uniteLibelle = _context.P_UNITE
                                     .Where(u => u.cbMarq == recup.AF_Unite)
                                     .Select(u => u.U_Intitule)
                                     .FirstOrDefault();

                                    gvLigneEdit.CustomColumnDisplayText += (s, e) =>
                                    {
                                        if (e.Column.FieldName == "Unite")
                                        {
                                            e.DisplayText = uniteLibelle ?? ""; // juste la string
                                        }
                                    };

                                    col1.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
                                    col1.SummaryItem.DisplayFormat = "{0:n2}";
                                    gvLigneEdit.Appearance.FooterPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                                }
                            }
                            catch (Exception pl) { }

                        }
                    }
                };

                GridColumn colFret = gvLigneEdit.Columns["DL_PoidsNet"];
                RepositoryItemTextEdit textEdit = new RepositoryItemTextEdit();
                colFret.ColumnEdit = textEdit;

                textEdit.KeyPress += (sender, e) =>
                {
                    // Remplacer le point par une virgule
                    if (e.KeyChar == '.')
                    {
                        e.KeyChar = ',';
                    }

                    // Autoriser uniquement chiffres, virgule et contrôle (backspace)
                    if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != ',')
                    {
                        e.Handled = true;  // bloque la saisie
                    }

                    // Autoriser une seule virgule
                    TextEdit editor = sender as TextEdit;
                    if (e.KeyChar == ',' && editor.Text.Contains(","))
                    {
                        e.Handled = true;
                    }
                };
            }
            else
            {
                txt_prix.EditValue = txt_poids.EditValue = "0";
            }

            gvLigneEdit.Columns["DL_PrixUnitaire"].Caption = "PU par Kg";
            RecalculerPrix(gvLigneEdit);
        }

        private void datelivrprev_EditValueChanged(object sender, EventArgs e)
        {
            datelivrprev.EditValue = datelivrprev.EditValue;
        }

        private void gvLigneEdit_CustomSummaryCalculate(object sender, DevExpress.Data.CustomSummaryEventArgs e)
        {
            GridView view = sender as GridView;

            if (e.SummaryProcess == CustomSummaryProcess.Start)
            {
                e.TotalValue = 0m;
            }

            if (e.SummaryProcess == CustomSummaryProcess.Calculate)
            {
                // Vérifie si "Retenu" est coché
                bool retenu = false;
                var valueRetenu = view.GetRowCellValue(e.RowHandle, "Retenu");
                if (valueRetenu != DBNull.Value && valueRetenu != null)
                    retenu = Convert.ToBoolean(valueRetenu);

                if (retenu)
                {
                    // FRAIS
                    if (e.IsTotalSummary && e.Item is GridSummaryItem summaryItemFrais &&
                        (summaryItemFrais.FieldName == "DL_Frais"))
                    {
                        var val = view.GetRowCellValue(e.RowHandle, "DL_Frais");
                        if (val != DBNull.Value && val != null)
                            e.TotalValue = (decimal)e.TotalValue + Convert.ToDecimal(val);
                    }
                    // MontantHT
                    if (e.IsTotalSummary && e.Item is GridSummaryItem summaryItem &&
                        (summaryItem.FieldName == "DL_MontantHT"))
                    {
                        var val = view.GetRowCellValue(e.RowHandle, "DL_MontantHT");
                        if (val != DBNull.Value && val != null)
                            e.TotalValue = (decimal)e.TotalValue + Convert.ToDecimal(val);
                    }


                    // MontantTTC
                    if (e.IsTotalSummary && e.Item is GridSummaryItem summaryItemTTC &&
                        (summaryItemTTC.FieldName == "DL_MontantTTC"))
                    {
                        var val = view.GetRowCellValue(e.RowHandle, "DL_MontantTTC");
                        if (val != DBNull.Value && val != null)
                            e.TotalValue = (decimal)e.TotalValue + Convert.ToDecimal(val);
                    }
                            
                    // MontantRegle
                    if (e.IsTotalSummary && e.Item is GridSummaryItem summaryItemRegle &&
                        (summaryItemRegle.FieldName == "DL_MontantRegle"))
                    {
                        var val = view.GetRowCellValue(e.RowHandle, "DL_MontantRegle");
                        if (val != DBNull.Value && val != null)
                            e.TotalValue = (decimal)e.TotalValue + Convert.ToDecimal(val);
                    }
                }
            }
        }

        private void gvLigneEdit_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView view = sender as GridView;

            if (e.RowHandle >= 0)
            {
                var value = view.GetRowCellValue(e.RowHandle, "Retenu");
                if (value != DBNull.Value && Convert.ToBoolean(value))
                {
                    e.Appearance.BackColor = Color.LightGreen; // couleur de fond
                    e.Appearance.ForeColor = Color.Black;      // couleur du texte (optionnel)
                }
            }
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ImprimerDocument();
        }
        private void ShowPreviewWithTracking(XtraReport report, string docPiece)
        {
            ReportPrintTool printTool = new ReportPrintTool(report);

            // Événement : impression
            printTool.PrintingSystem.EndPrint += (s, e) =>
            {
                MarquerCommeImprime(docPiece);
            };

            //// Événement : export
            //printTool.PrintingSystem. += (s, e) =>
            //{
            //    MarquerCommeImprime(docPiece);
            //};

            printTool.ShowPreviewDialog();
        }
        private void MarquerCommeImprime(string docPiece)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = "UPDATE F_DOCENTETE SET DO_Imprim = 1 WHERE DO_Piece = @DocPiece";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@DocPiece", docPiece);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tester_cloturer(dopiecetxt.Text))
            {
                MessageBox.Show(
                    "Ce document est déjà clôturé, vous ne pouvez plus modifier son contenu!!!!",
                    "Modification bloquée",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            else
            {
                string connectionStringArbio = $"Server={serveripPrincipale};" +
                                $"Database=TRANSIT;User ID=Dev;Password=1234;" +
                                $"TrustServerCertificate=True;Connection Timeout=120;";

                using (SqlConnection connection = new SqlConnection(connectionStringArbio))
                {
                    connection.Open();

                    using (SqlTransaction tran = connection.BeginTransaction())
                    {
                        try
                        {
                            string sql = @"
                            DELETE FROM encoursutilisation
                            WHERE utilisateur = @utilisateur AND numero_doc=@numero_doc";

                            using (SqlCommand cmd = new SqlCommand(sql, connection, tran))
                            {
                                cmd.Parameters.Add("@utilisateur", SqlDbType.VarChar)
                                              .Value = FrmMdiParent.IDName;
                                cmd.Parameters.Add("@numero_doc", SqlDbType.VarChar)
                                             .Value = dopiecetxt.Text;

                                cmd.ExecuteNonQuery();
                            }

                            // Validation
                            tran.Commit();
                        }
                        catch (Exception ex)
                        {
                            // Annulation si erreur
                            if (tran != null)
                                tran.Rollback();

                            MessageBox.Show(
                                ex.Message,
                                "Erreur",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                }

                if (StatutActuel < 2)
                {
                    MessageBox.Show("Le statut actuel ne permet pas de transformer le document.", "Information");
                    return;
                }

                if (dopiecetxt.Text.ToString().StartsWith("APA"))
                {
                    bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "TRANSFORM");

                    if (autorise)
                    {
                        if(lkStatut.Text == "Accepté") {
                            var dlg = new frmTransform(_typeDocument);
                            dlg.ParentFormInstance = this;

                            if (dlg.ShowDialog() == DialogResult.OK)
                            {
                                this.TransformFDOCENTETE(dlg.doctype);
                            }

                            // Fix for CS0120: Use the instance of ucDocuments instead of trying to call it statically
                            _ucDocuments.ChargerDonneesDepuisBDD();
                        }
                        else
                        {
                            MessageBox.Show(
                                   "Le statut actuel ne vous permet pas de transformer ce document !",
                                   "Modification bloquée",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show(
                            "Vous n'avez pas l'autorisation de transformer un projet d'achat !",
                            "Transformation bloquée",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                }
                else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
                {
                    bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "TRANSFORM");

                    if (autorise)
                    {
                        var dlg = new frmTransform(_typeDocument);
                        dlg.ParentFormInstance = this;

                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            this.TransformFDOCENTETE(dlg.doctype);
                        }

                        // Fix for CS0120: Use the instance of ucDocuments instead of trying to call it statically
                        _ucDocuments.ChargerDonneesDepuisBDD();
                    }
                    else
                    {
                        MessageBox.Show(
                            "Vous n'avez pas l'autorisation de transformer un bon de réception !",
                            "Transformation bloquée",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                }
                else if (dopiecetxt.Text.ToString().StartsWith("AFA"))
                {
                    bool autorise = frmMenuAchat.verifier_droit("Facture", "TRANSFORM");

                    if (autorise)
                    {
                        var dlg = new frmTransform(_typeDocument);
                        dlg.ParentFormInstance = this;

                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            this.TransformFDOCENTETE(dlg.doctype);
                        }

                        // Fix for CS0120: Use the instance of ucDocuments instead of trying to call it statically
                        _ucDocuments.ChargerDonneesDepuisBDD();
                    }
                    else
                    {
                        MessageBox.Show(
                            "Vous n'avez pas l'autorisation de transformer une facture !",
                            "Transformation bloquée",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                }
                else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
                {
                    bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "TRANSFORM");

                    if (autorise)
                    {
                        var dlg = new frmTransform(_typeDocument);
                        dlg.ParentFormInstance = this;

                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            this.TransformFDOCENTETE(dlg.doctype);
                        }

                        //MAJ Qte

                        /*int? DE_No = Convert.ToInt32(lkDepot.EditValue);
                        for (int i = 0; i < gvLigneEdit.RowCount; i++)
                        {
                            string reference = gvLigneEdit.GetRowCellValue(i, "AR_Ref")?.ToString();
                            var qte = gvLigneEdit.GetRowCellValue(i, "DL_Qte");

                            try
                            {
                                 string connectionString = $"Server={serveripPrincipale};Database=TRANSIT;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";

                                 using (SqlConnection connection = new SqlConnection(connectionString))
                                 {
                                     connection.Open();

                                     string existingStock = @"
                                     SELECT TOP 1 AS_QteSto
                                     FROM F_ARTSTOCK
                                     WHERE AR_Ref = @AR_Ref AND DE_No = @DE_No";

                                     decimal? existingQte = null;

                                     using (SqlCommand checkCmd = new SqlCommand(existingStock, connection))
                                     {
                                         checkCmd.Parameters.AddWithValue("@AR_Ref", reference);
                                         checkCmd.Parameters.AddWithValue("@DE_No", DE_No);

                                         var result = checkCmd.ExecuteScalar();
                                         if (result != null && result != DBNull.Value)
                                             existingQte = Convert.ToDecimal(result);
                                     }
                                     if (existingQte == null)
                                     {
                                         string insertSql = @"
                                             INSERT INTO F_ARTSTOCK (
                                                 AR_Ref,
                                                 DE_No,
                                                 DP_NoPrincipal,
                                                 AS_QteSto,
                                                 AS_QteRes,
                                                 AS_QteCom,
                                                 AS_QtePrepa,
                                                 AS_MontSto,
                                                 AS_QteMini,
                                                 AS_QteMaxi
                                             )
                                             VALUES (
                                                 @AR_Ref,
                                                 @DE_No,
                                                 @DP_NoPrincipal,
                                                 @Qte,
                                                 0, 0, 0, 0, 0, 0
                                             )";

                                         using (SqlCommand insertCmd = new SqlCommand(insertSql, connection))
                                         {
                                             insertCmd.Parameters.AddWithValue("@AR_Ref", reference);
                                             insertCmd.Parameters.AddWithValue("@DE_No", DE_No);
                                             insertCmd.Parameters.AddWithValue("@DP_NoPrincipal", 1);
                                             insertCmd.Parameters.AddWithValue("@Qte", qte);

                                             insertCmd.ExecuteNonQuery();
                                         }
                                     }
                                     else
                                     {
                                             string updateSql = @"
                                             UPDATE F_ARTSTOCK
                                             SET AS_QteSto = AS_QteSto + @Qte
                                             WHERE AR_Ref = @AR_Ref AND DE_No = @DE_No";

                                             using (SqlCommand updateCmd = new SqlCommand(updateSql, connection))
                                             {
                                                 updateCmd.Parameters.AddWithValue("@AR_Ref", reference);
                                                 updateCmd.Parameters.AddWithValue("@DE_No", DE_No);
                                                 updateCmd.Parameters.AddWithValue("@Qte", qte);

                                                 updateCmd.ExecuteNonQuery();
                                             }
                                     }
                                 }

                                //MAJ des articles par lot
                                simpleButton2_Click(sender, e);   
                            }
                            catch (System.Data.Entity.Infrastructure.DbUpdateException ex)
                            {
                               
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Erreur");
                            }
                               
                            
                        }*/
                        // Fix for CS0120: Use the instance of ucDocuments instead of trying to call it statically
                        _ucDocuments.ChargerDonneesDepuisBDD();
                    }
                    else
                    {
                        MessageBox.Show(
                            "Vous n'avez pas l'autorisation de transformer un bon de commande!",
                            "Transformation bloquée",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                }
            }
        }

        private void gvLigneEdit_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            if (e.Row != null && gvLigneEdit.IsNewItemRow(e.RowHandle))
            {
                AddLigne();
                MettreAJourTotauxDepuisBD(dopiecetxt.Text);
            }
        }

        private void xtraTabControl1_Selected(object sender, DevExpress.XtraTab.TabPageEventArgs e)
        {
            chkActif.Checked = true;

        }
        private int nbchecked = 0;
        private void chkActif_CheckedChanged(object sender, EventArgs e)
        {
            if (dopiecetxt.Text.ToString().StartsWith("AFA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour une facture !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("APA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un projet d'achat !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un bon de commande !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un Bon de réception !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private void chkSommeil_CheckedChanged(object sender, EventArgs e)
        {
            if (dopiecetxt.Text.ToString().StartsWith("AFA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour une facture !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("APA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un projet d'achat !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un bon de commande !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                if (autorise)
                {
                    MettreAJourFiltre();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un Bon de réception !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }
        public System.Windows.Forms.BindingSource BindingFrns = new System.Windows.Forms.BindingSource();
        public System.Windows.Forms.BindingSource BindingArt = new System.Windows.Forms.BindingSource();
        private void MettreAJourFiltre()
        {
            bool actif = chkActif.Checked;
            bool sommeil = chkSommeil.Checked;

            Entetes.FiltrerFournisseurs(gcFrns, actif, sommeil, BindingFrns);
            RepositoryItemHyperLinkEdit hyperlink = new RepositoryItemHyperLinkEdit();
            hyperlink.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            hyperlink.Click += HyperlinkFrns_Click;
            gcFrns.RepositoryItems.Add(hyperlink);
            gvFrns.Columns["CT_Num"].ColumnEdit = hyperlink;
            foreach (GridColumn column in gvFrns.Columns)
            {
                if (column.FieldName == "CT_Num" || column.FieldName == "CT_Intitule")
                    continue; // Ne pas masquer la colonne CT_Num   
                column.VisibleIndex = -1;
            }


        }
        private void showFrmEditFrns(string ctNum)
        {
            try
            {
                if (!string.IsNullOrEmpty(ctNum))
                {
                    frmEditFrns frm = new frmEditFrns(ctNum, BindingFrns);
                    frm.ShowDialog();
                    frm.DonneesMisesAJour += (s, args) =>
                    {
                        // Appel de ta méthode locale
                        MettreAJourFiltre();
                    };
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void HyperlinkFrns_Click(object sender, EventArgs e)
        {
            int rowHandle = gvFrns.FocusedRowHandle;
            if (rowHandle >= 0)
            {
                string ctNum = gvFrns.GetRowCellValue(rowHandle, "CT_Num").ToString();
                if (rowHandle >= 0)
                {
                    showFrmEditFrns(ctNum);
                }
            }

        }

        private void hyperlinkLabelControl6_Click(object sender, EventArgs e)
        {
            MettreAJourFiltre();
        }
        //P_DEVISERepository _p_DEVISERepository;
        //private void lkEdFrns_EditValueChanged(object sender, EventArgs e)
        //{
        //    if (lkEdFrns.EditValue == null)
        //    {
        //        lkDevise.EditValue = null;
        //        return;
        //    }

        //    string ctNum = lkEdFrns.EditValue.ToString();

        //    F_COMPTET frns = _f_COMPTETRepository?.GetDeviseNameByCT_Num(ctNum);

        //    if (frns != null)
        //    {
        //        P_DEVISE devise = _p_DEVISERepository?.Get_P_DEVISE_By_cbMarq(frns.N_Devise ?? 0);
        //        lkDevise.EditValue = devise?.D_Intitule;
        //    }
        //    else
        //    {
        //        lkDevise.EditValue = null;
        //    }

        //}

        private void hyperlinkLabelControl7_Click(object sender, EventArgs e)
        {
            if(lkEdFrns.EditValue != null)
            {
                string ctNum = lkEdFrns?.EditValue.ToString();
                Entetes.FiltrerFournisseurs(gcFrns, true, false, BindingFrns);
                int rownum = BindingFrns.Count;
                showFrmEditFrns(ctNum);
            }
            
        }
        
        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                decimal totalTVA = 0;
                List<F_DOCLIGNE> lignes = _f_DOCLIGNERepository.GetAll_F_DOCLIGNE_Of_DOCENTETE(dopiece);
                decimal totalHT = lignes.Sum(l => l.DL_MontantHT ?? 0);

                // appliquer 20% de TVA
                totalTVA = totalHT * 0.20m;

                bool toutesLesLignesValides = true;
                for (int i = 0; i < gvLigneEdit.RowCount; i++)
                {
                    int rowHandle = gvLigneEdit.GetVisibleRowHandle(i);
                    if (gvLigneEdit.IsDataRow(rowHandle))
                    {
                        var laligneamettreajour = gvLigneEdit.GetRow(rowHandle);

                        // Accès à la cellule "DL_PrixUnitaire"
                        object cellValue = gvLigneEdit.GetRowCellValue(rowHandle, "DL_PrixUnitaire");

                        // Vérifie que la cellule contient une valeur convertible en décimal
                        if (cellValue != null && decimal.TryParse(cellValue.ToString(), out decimal puBrut))
                        {
                            if (puBrut == 0)
                            {
                                DevExpress.XtraEditors.XtraMessageBox.Show(
                                    "Veuillez renseigner le PU dans toutes les lignes et faire un Update par ligne pour que le calcul de frais se fera dans la logique",
                                    "Erreur",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                                toutesLesLignesValides = false;
                                break;
                            }
                        }
                    }
                }

                if (toutesLesLignesValides)
                {
                    frmEditInfos frm = new frmEditInfos(dopiece, totalTVA, this);
                    frm.ShowDialog();
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        ExcelDataSource excelDataSource;
        private void hyperlinkLabelControl5_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Tous les fichiers (*.*)|*.*|Images (*.jpg;*.png)|*.jpg;*.png";
                    openFileDialog.Title = "Sélectionner un fichier";
                    openFileDialog.Multiselect = false;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;
                        if (!File.Exists(filePath))
                        {
                            XtraMessageBox.Show("Le fichier sélectionné n'existe pas.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Charger les données depuis Excel
                        Workbook workbook = new Workbook();
                        workbook.LoadDocument(filePath);

                        if (workbook.Worksheets.Count == 0)
                        {
                            XtraMessageBox.Show("Le fichier Excel ne contient aucune feuille.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        string sheetName = workbook.Worksheets[0].Name;

                        ExcelDataSource excelDataSource = new ExcelDataSource
                        {
                            FileName = filePath,
                            SourceOptions = new ExcelSourceOptions(new ExcelWorksheetSettings(sheetName, "A:E"))
                        };

                        excelDataSource.Schema.AddRange(new FieldInfo[]
                        {
                            new FieldInfo { Name = "CT_Num", Type = typeof(string) },
                            new FieldInfo { Name = "AR_Ref", Type = typeof(string) },
                            new FieldInfo { Name = "DL_Design", Type = typeof(string) },
                            new FieldInfo { Name = "DL_PrixUnitaire", Type = typeof(decimal) },
                            new FieldInfo { Name = "DL_Qte", Type = typeof(decimal) }
                        });

                        excelDataSource.Fill();
                        DataTable dataDocLigneImport = excelDataSource.ToDataTable();

                        if (dataDocLigneImport == null || dataDocLigneImport.Rows.Count == 0)
                        {
                            XtraMessageBox.Show("Aucune ligne trouvée dans le fichier Excel.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        // Validation des données avant traitement
                        List<string> erreursValidation = new List<string>();
                        for (int i = 0; i < dataDocLigneImport.Rows.Count; i++)
                        {
                            DataRow row = dataDocLigneImport.Rows[i];

                            if (string.IsNullOrWhiteSpace(row["AR_Ref"]?.ToString()))
                            {
                                erreursValidation.Add($"Ligne {i + 2}: AR_Ref manquant");
                            }
                            if (string.IsNullOrWhiteSpace(row["CT_Num"]?.ToString()))
                            {
                                erreursValidation.Add($"Ligne {i + 2}: CT_Num manquant");
                            }
                            if (row["DL_Qte"] == DBNull.Value || Convert.ToDecimal(row["DL_Qte"]) <= 0)
                            {
                                erreursValidation.Add($"Ligne {i + 2}: Quantité invalide");
                            }
                        }

                        if (erreursValidation.Count > 0)
                        {
                            string message = "Erreurs de validation détectées:\n" + string.Join("\n", erreursValidation.Take(10));
                            if (erreursValidation.Count > 10)
                            {
                                message += $"\n... et {erreursValidation.Count - 10} autres erreurs";
                            }
                            XtraMessageBox.Show(message, "Erreurs de validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        string rec_arref = "";
                        bool t_ok = false;
                        int lignesTraitees = 0;
                        int lignesErreur = 0;

                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();

                            int total = dataDocLigneImport.Rows.Count;
                            int current = 0;

                            foreach (DataRow row in dataDocLigneImport.Rows)
                            {
                                try
                                {
                                    int dotype = 10;
                                    string arRef = row["AR_Ref"]?.ToString()?.Trim() ?? "";
                                    string ctNum = row["CT_Num"]?.ToString()?.Trim() ?? "";
                                    rec_arref = arRef;

                                    decimal dlQte = row["DL_Qte"] != DBNull.Value ? Convert.ToDecimal(row["DL_Qte"]) : 0;
                                    decimal DLTaxe1 = 0;

                                    // Vérifier l'existence de la référence
                                    if (not_existence_ref(arRef))
                                    {
                                        lignesErreur++;
                                        XtraMessageBox.Show($"La référence : {arRef} n'existe pas dans la base, veuillez vous rapprocher du service informatique!",
                                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        break;
                                    }

                                    current++;
                                    SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
                                    SplashScreenManager.Default.SetWaitFormDescription($"Traitement: {current}/{total}");

                                    F_ARTICLE articleChoisi = _f_ARTICLERepository.GetF_ARTICLEByAR_Ref(arRef);

                                    if (articleChoisi == null)
                                    {
                                        lignesErreur++;
                                        SplashScreenManager.CloseForm();
                                        XtraMessageBox.Show($"Article introuvable pour la référence : {arRef}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        break;
                                    }

                                    int? DE_No = Convert.ToInt32(lkDepot.EditValue);
                                    decimal remisePourcent = 0;
                                    string arDesign = row["DL_Design"]?.ToString()?.Trim() ?? "";
                                    decimal puBrut = row["DL_PrixUnitaire"] != DBNull.Value ? Convert.ToDecimal(row["DL_PrixUnitaire"]) : 0;
                                    decimal puNet = puBrut * (1 - remisePourcent / 100);
                                    decimal montantHT = dlQte * puNet;
                                    decimal montantTTC = montantHT * (1 + DLTaxe1 / 100);
                                    int retenu = 1;

                                    F_DOCENTETE docEnCours = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(dopiecetxt.Text);

                                    if (docEnCours == null)
                                    {
                                        lignesErreur++;
                                        SplashScreenManager.CloseForm();
                                        XtraMessageBox.Show($"Document en-tête introuvable : {dopiecetxt.Text}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        break;
                                    }

                                    F_COLLABORATEUR collab = _listeCollaborateurs?.FirstOrDefault(c => c.CO_No == (int)lkEdCollaborateur.EditValue);

                                    DateTime DO_Date = dateSaisie.DateTime;
                                    DateTime DO_DateLivr = (datelivrprev.EditValue is DateTime)
                                        ? (DateTime)datelivrprev.EditValue
                                        : new DateTime(1753, 01, 01);

                                    // Calcul du numéro de ligne
                                    int? maxValueDL_Ligne = 0;
                                    for (int i = 0; i < gvLigneEdit.RowCount; i++)
                                    {
                                        object dlligne = gvLigneEdit.GetRowCellValue(i, "DL_Ligne");
                                        if (dlligne != null && int.TryParse(dlligne.ToString(), out int val))
                                        {
                                            if (val > maxValueDL_Ligne)
                                            {
                                                maxValueDL_Ligne = val;
                                            }
                                        }
                                    }

                                    int? numeroLigneDL_Ligne = maxValueDL_Ligne + 1000;
                                    string reference = txtDoRef.Text;
                                    short typeDoc = (short)_f_DOCENTETEService.GetDocTypeNo(_prefix);
                                    short DL_NoRef = (short)(gvLigneEdit.RowCount + current);
                                    int dlno = GetMaxDLNo() + current;

                                    _f_DOCLIGNEService.AjouterF_DOCLIGNE(
                                        typeDoc,
                                        ctNum,
                                        dopiecetxt.Text,
                                        DO_Date,
                                        numeroLigneDL_Ligne,
                                        docEnCours,
                                        arRef,
                                        arDesign,
                                        DLTaxe1,
                                        dlQte,
                                        typeDoc.ToString(),
                                        articleChoisi,
                                        Convert.ToString(dlQte),
                                        remisePourcent.ToString(),
                                        puNet.ToString(),
                                        collab,
                                        DL_NoRef,
                                        puBrut,
                                        DO_DateLivr,
                                        comboBoxAffaire.Text,
                                        montantTTC.ToString(),
                                        montantHT.ToString(),
                                        DateTime.Now,
                                        DE_No,
                                        dlno,
                                        retenu);

                                    lignesTraitees++;
                                    t_ok = true;
                                }
                                catch (Exception exRow)
                                {
                                    lignesErreur++;
                                    SplashScreenManager.CloseForm();
                                    XtraMessageBox.Show($"Erreur lors du traitement de la ligne {current}:\n{exRow.Message}",
                                        "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                            }

                            // Mise à jour finale
                            if (t_ok && lignesTraitees > 0)
                            {
                                try
                                {
                                    MettreAJourTotauxDepuisBD(dopiecetxt.Text);
                                    string _currentDocPieceNo = dopiecetxt.Text;
                                    InitializeGrid(gcLigneEdit, _currentDocPieceNo);

                                    SplashScreenManager.CloseForm();
                                    XtraMessageBox.Show($"Importation terminée avec succès.\n{lignesTraitees} ligne(s) importée(s).",
                                        "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception exFinal)
                                {
                                    SplashScreenManager.CloseForm();
                                    XtraMessageBox.Show($"Erreur lors de la mise à jour finale:\n{exFinal.Message}",
                                        "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else if (lignesErreur > 0)
                            {
                                SplashScreenManager.CloseForm();
                                XtraMessageBox.Show($"Importation interrompue. {lignesTraitees} ligne(s) traitée(s), {lignesErreur} erreur(s).",
                                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                SplashScreenManager.CloseForm();
                MethodBase m = MethodBase.GetCurrentMethod();
                XtraMessageBox.Show($"Une erreur est survenue:\n{ex.Message}\n\nMéthode: {m.Name}\n\nStack Trace:\n{ex.StackTrace}",
                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void MettreAJourTotauxDepuisBD(string doPiece)
        {
            string query = @"
                            SELECT 
                                ISNULL(SUM(DL_MontantHT), 0) AS TotalHT,
                                ISNULL(SUM(DL_MontantTTC), 0) AS TotalTTC
                            FROM F_DOCLIGNE
                            WHERE DO_Piece = @DOPiece AND Retenu <> 0";

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@DOPiece", doPiece);

                try
                {
                    conn.Open();

                    decimal totalHT = 0, totalTTC = 0, netAPayer = 0;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            totalHT = reader.GetDecimal(0);
                            totalTTC = reader.GetDecimal(1);
                        }
                    }

                    string updateQuery = @"
                                    UPDATE F_DOCENTETE
                                    SET 
                                        DO_TotalHT = @TotalHT,
                                        DO_TotalTTC = @TotalTTC,
                                        DO_NetAPayer = @NetAPayer
                                    WHERE DO_Piece = @DOPiece";

                    using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                    {
                        updateCmd.Parameters.AddWithValue("@TotalHT", totalHT);
                        updateCmd.Parameters.AddWithValue("@TotalTTC", totalTTC);
                        updateCmd.Parameters.AddWithValue("@NetAPayer", totalTTC);
                        updateCmd.Parameters.AddWithValue("@DOPiece", doPiece);

                        updateCmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void CustomLayout()
        {
            switch (_prefix)
            {
                case "AFA":
                    lciTreelist.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    break;

                case "ABL":
                    lciTreelist.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    break;

                case "ABC":
                    lciTreelist.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    break;

                case "ABR":
                    lciTreelist.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    break;
            }
        }

        private void barButtonItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            AccesDocLie();
            LoadDocLie(dopiece);
        }
        private void AccesDocLie()
        {
            try
            {
                using (XtraOpenFileDialog openFileDialog = new XtraOpenFileDialog())
                {
                    openFileDialog.Title = "Sélectionnez les fichiers à transférer";
                    openFileDialog.Multiselect = true; // ✅ Autoriser plusieurs fichiers
                    openFileDialog.Filter = "Tous les fichiers (*.*)|*.*";

                    // 2. Ouvrir la boîte de dialogue
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // 3. Récupérer la liste des fichiers sélectionnés
                        string[] selectedFiles = openFileDialog.FileNames;

                        // 4. Définir le dossier réseau de destination
                        string destinationFolder = @"\\Srv-arb\documents_achats$";
                        string nodoc = dopiece.Substring(3, 8);

                        if (dopiecetxt.Text.ToString().StartsWith("ABC"))
                        {
                            bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                            if (autorise)
                            {
                                if (tester_cloturer(dopiecetxt.Text))
                                {
                                    MessageBox.Show(
                                        "Ce document est déjà clôturé, vous ne pouvez plus modifier son contenu!!!!",
                                        "Modification bloquée",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error
                                    );
                                }
                                else
                                {
                                    try
                                    {
                                        string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";
                                        if (!Directory.Exists(destinationFolderdoc))
                                        {
                                            Directory.CreateDirectory(destinationFolderdoc);
                                        }

                                        // 5. Copier chaque fichier
                                        foreach (string file in selectedFiles)
                                        {
                                            string fileName = Path.GetFileName(file);
                                            string destFile = Path.Combine(destinationFolderdoc, fileName);

                                            File.Copy(file, destFile, true); // true = overwrite si existe déjà
                                        }

                                        XtraMessageBox.Show("✅ Transfert terminé avec succès !", "Succès",
                                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                    catch (Exception ex)
                                    {
                                        MethodBase m = MethodBase.GetCurrentMethod();
                                        MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show(
                                    "Vous n'avez pas l'autorisation de modifier un bon de commande!",
                                    "Modification bloquée",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                            }
                        }
                        else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
                        {
                            bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                            if (autorise)
                            {
                                try
                                {
                                    string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";
                                    if (!Directory.Exists(destinationFolderdoc))
                                    {
                                        Directory.CreateDirectory(destinationFolderdoc);
                                    }

                                    // 5. Copier chaque fichier
                                    foreach (string file in selectedFiles)
                                    {
                                        string fileName = Path.GetFileName(file);
                                        string destFile = Path.Combine(destinationFolderdoc, fileName);

                                        File.Copy(file, destFile, true); // true = overwrite si existe déjà
                                    }

                                    XtraMessageBox.Show("✅ Transfert terminé avec succès !", "Succès",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MethodBase m = MethodBase.GetCurrentMethod();
                                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show(
                                    "Vous n'avez pas l'autorisation de modifier un bon de réception!",
                                    "Modification bloquée",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                            }
                        }
                        else if (dopiecetxt.Text.ToString().StartsWith("AFA"))
                        {
                            bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                            if (autorise)
                            {
                                try
                                {
                                    string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";
                                    if (!Directory.Exists(destinationFolderdoc))
                                    {
                                        Directory.CreateDirectory(destinationFolderdoc);
                                    }

                                    // 5. Copier chaque fichier
                                    foreach (string file in selectedFiles)
                                    {
                                        string fileName = Path.GetFileName(file);
                                        string destFile = Path.Combine(destinationFolderdoc, fileName);

                                        File.Copy(file, destFile, true); // true = overwrite si existe déjà
                                    }

                                    XtraMessageBox.Show("✅ Transfert terminé avec succès !", "Succès",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MethodBase m = MethodBase.GetCurrentMethod();
                                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show(
                                    "Vous n'avez pas l'autorisation de modifier une facture!",
                                    "Modification bloquée",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                            }
                        }
                        else if (dopiecetxt.Text.ToString().StartsWith("APA"))
                        {
                            bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                            if (autorise)
                            {
                                try
                                {
                                    string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";
                                    if (!Directory.Exists(destinationFolderdoc))
                                    {
                                        Directory.CreateDirectory(destinationFolderdoc);
                                    }

                                    // 5. Copier chaque fichier
                                    foreach (string file in selectedFiles)
                                    {
                                        string fileName = Path.GetFileName(file);
                                        string destFile = Path.Combine(destinationFolderdoc, fileName);

                                        File.Copy(file, destFile, true); // true = overwrite si existe déjà
                                    }

                                    XtraMessageBox.Show("✅ Transfert terminé avec succès !", "Succès",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MethodBase m = MethodBase.GetCurrentMethod();
                                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show(
                                    "Vous n'avez pas l'autorisation de modifier un Projet d'achat!",
                                    "Modification bloquée",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                            }
                        }
                    }
                }
            }

            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void LoadDocLie(string DoPiece)
        {
            try
            {
                listView1.Items.Clear();
                imageList1.Images.Clear();

                listView1.View = System.Windows.Forms.View.Details;
                listView1.FullRowSelect = true;
                listView1.Columns.Add("Nom", 200);
                listView1.Columns.Add("Taille", 100);
                //listView1.Columns.Add("Type", 150);
                listView1.Columns.Add("Date de modification", 150);

                imageList1 = new ImageList();
                imageList1.ColorDepth = ColorDepth.Depth32Bit;
                imageList1.ImageSize = new Size(16, 16);

                listView1.SmallImageList = imageList1;

                string nodoc = DoPiece.Substring(3, 8);
                string destinationFolderdoc = $@"\\Srv-arb\documents_achats$\{nodoc}";

                // Vérifie si le dossier existe, sinon le créer
                if (!Directory.Exists(destinationFolderdoc))
                {
                    Directory.CreateDirectory(destinationFolderdoc);
                }

                var dirInfo = new DirectoryInfo(destinationFolderdoc);

                //// Dossiers
                //foreach (var dir in dirInfo.GetDirectories())
                //{
                //    int iconIndex = AddIconToImageList(dir.FullName);
                //    var item = new ListViewItem(dir.Name, iconIndex);
                //    item.SubItems.Add("");
                //    item.SubItems.Add("Dossier");
                //    item.SubItems.Add(dir.LastWriteTime.ToString());
                //    item.Tag = dir.FullName;
                //    listView1.Items.Add(item);
                //}

                // Fichiers
                foreach (var file in dirInfo.GetFiles())
                {
                    Icon icon = FileIconHelper.GetSmallIcon(file.FullName);
                    imageList1.Images.Add(file.Extension, icon);

                    ListViewItem item = new ListViewItem(file.Name);
                    item.ImageKey = file.Extension;
                    item.SubItems.Add(file.Length.ToString());
                    //item.SubItems.Add(file.Extension);
                    item.SubItems.Add(file.LastWriteTime.ToString());

                    listView1.Items.Add(item);
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private int AddIconToImageList(string path)
        {
            Icon icon = Icon.ExtractAssociatedIcon(path);
            if (icon != null)
            {
                imageList1.Images.Add(icon);
                return imageList1.Images.Count - 1;
            }
            return -1;
        }
        private int GetFileIconIndex(string filePath)
        {
            Icon icon = Icon.ExtractAssociatedIcon(filePath);
            if (icon != null)
            {
                imageList1.Images.Add(icon);
                return imageList1.Images.Count - 1;
            }
            return -1;
        }
        private void ribbonControl1_Click(object sender, EventArgs e)
        {

        }

        private void lkStatut_EditValueChanging(object sender, ChangingEventArgs e)
        {
            if (lkStatut.Text != "Accepté")
            {
                if (e.NewValue == null) return;
                int newValue = Convert.ToInt32(e.NewValue);

                if (newValue < StatutActuel)
                {
                    e.Cancel = true;
                    MessageBox.Show("Impossible de revenir à un statut précédent.", "Action non autorisée", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void lkStatut_EditValueChanged(object sender, EventArgs e)
        {
            if (lkStatut.EditValue != null )
            {
                StatutActuel = lkStatut.EditValue == null || lkStatut.EditValue == DBNull.Value ? 0 : Convert.ToInt32(lkStatut.EditValue);
            }
        }

        private void maps1_MouseMove(object sender, MouseEventArgs e)
        {
            PointF point = new PointF(e.X, e.Y);

            // Fix: Replace the non-existent PointToLatLng method with GetMapCoordinates
           // PointF latLong = this.maps1.GetMapCoordinates(point);

           // float latitude = latLong.Y;
           // float longitude = latLong.X;

            // Affiche dans la barre de statut ou un Label
           // lblCoord.Text = $"Lat: {latitude:F4}, Lon: {longitude:F4}";

            // Ou dans un tooltip
            System.Windows.Forms.ToolTip tt = new System.Windows.Forms.ToolTip();
            //tt.SetToolTip(maps1, $"Lat: {latitude:F4}, Lon: {longitude:F4}");
        }
        private void Hyperlink_ClickArticle(object sender, EventArgs e)
        {
            int rowHandle = gvArticle.FocusedRowHandle;
            if (rowHandle >= 0)
            {
                string arRef = gvArticle.GetRowCellValue(rowHandle, "AR_Ref").ToString();
                if (rowHandle >= 0)
                {
                    frmEditArticles frm = new frmEditArticles(arRef, BindingArt);
                    frm.ShowDialog();
                }
            }
        }
        

        private void FiltrerArticles()
        {
            // On s'assure de cibler le GridView
            GridView view = gcArticle.MainView as GridView;

            string filter = "";

            if (chkArActif.Checked && !chkArSommeil.Checked)
                filter = "[AR_Sommeil] = 0";

            else if (!chkArActif.Checked && chkArSommeil.Checked)
                filter = "[AR_Sommeil] = 1";

            else if (chkArActif.Checked && chkArSommeil.Checked)
                filter = ""; // tout afficher

            view.ActiveFilterString = filter;
        }


        private void chkArActif_CheckedChanged(object sender, EventArgs e)
        {
            if (dopiecetxt.Text.ToString().StartsWith("AFA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour une facture !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("APA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un projet d'achat !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un bon de commande !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un Bon de réception !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private void chkArSommeil_CheckedChanged(object sender, EventArgs e)
        {
            if (dopiecetxt.Text.ToString().StartsWith("AFA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour une facture !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );

                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("APA"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un projet d'achat !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABC"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un bon de commande !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
            else if (dopiecetxt.Text.ToString().StartsWith("ABR"))
            {
                bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                if (autorise)
                {
                    FiltrerArticles();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de mettre à jour un Bon de réception !",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private void hyperlinkLabelControl8_Click(object sender, EventArgs e)
        {
            ChargerArticles();
        }

        private void frmEditDocument_Activated(object sender, EventArgs e)
        {
            bindingNavigator1.Visible=false;
            if (dopiecetxt.Text.StartsWith("AFA") || dopiecetxt.Text.StartsWith("ABR") || dopiecetxt.Text.StartsWith("APC"))
            {
                gvLigneEdit.OptionsBehavior.Editable = false;
            }
            else
            {
                gvLigneEdit.OptionsBehavior.Editable = true;
            }

            // === UNITE ===
            GridColumn col_unite = gvLigneEdit.Columns["Unite"];
            if (col_unite == null)
            {
                col_unite = gvLigneEdit.Columns.AddField("Unite");
                col_unite.Caption = "Unité";
                col_unite.UnboundType = DevExpress.Data.UnboundColumnType.Decimal;
                col_unite.Visible = true;
                col_unite.OptionsColumn.AllowEdit = false;
                col_unite.OptionsColumn.ReadOnly = true;
                col_unite.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            }

            // === FRET ===
            GridColumn col = gvLigneEdit.Columns["FRET"];
            if (col == null)
            {
                col = gvLigneEdit.Columns.AddField("FRET");
                col.Caption = "FRET";
                col.UnboundType = DevExpress.Data.UnboundColumnType.Decimal;
                col.Visible = true;
                col.OptionsColumn.AllowEdit = true;
                col.OptionsColumn.ReadOnly = false;
            }

            // === POSITIONNEMENT ===
            int indexFrais = gvLigneEdit.Columns["DL_Frais"].VisibleIndex;
            col.VisibleIndex = indexFrais;

            int indexPrixUnitaire = gvLigneEdit.Columns["DL_PrixUnitaire"].VisibleIndex;
            col_unite.VisibleIndex = indexPrixUnitaire - 1;
        }

        private void RecalculerPrix(GridView view)
        {
            for (int i = 0; i < view.RowCount; i++)
            {
                if (view.IsGroupRow(i)) continue;

                decimal prixKg = 0;
                decimal prixTonne = 0;

                object objKg = view.GetRowCellValue(i, "DL_PrixUnitaire");
                object objTonne = view.GetRowCellValue(i, "PU par tonne");

                if (decimal.TryParse(objKg?.ToString(), out prixKg))
                {
                    // kg → tonne
                    prixTonne = prixKg * 1000;
                    view.SetRowCellValue(i, "PU par tonne", prixTonne);
                }
                else if (decimal.TryParse(objTonne?.ToString(), out prixTonne))
                {
                    // tonne → kg
                    prixKg = prixTonne / 1000;
                    view.SetRowCellValue(i, "DL_PrixUnitaire", prixKg);
                }
            }
        }

        private void lkDevise_EditValueChanged_1(object sender, EventArgs e)
        {
            var view = gvLigneEdit;
            if (view.RowCount <= 0) return;

            var colQte = view.Columns["DL_QTE"];
            var colPU = view.Columns["DL_PrixUnitaire"];
            var colMontantHT = view.Columns["DL_montantHT"];

            if (colQte == null || colPU == null ||  colMontantHT == null) return;

            for (int i = 0; i < view.RowCount; i++)
            {
                // Récupération des valeurs
                decimal qte = Convert.ToDecimal(view.GetRowCellValue(i, colQte));
                decimal pu = Convert.ToDecimal(view.GetRowCellValue(i, colPU));
                decimal cours = Convert.ToDecimal(txtCours.Text);

                // Calcul du montant HT
                decimal montantHT = qte * pu * cours;

                // Mise à jour de la cellule
                view.SetRowCellValue(i, colMontantHT, montantHT);

                // Déclenchement du handler si nécessaire
                var args = new DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs(
                    i,
                    colMontantHT,
                    montantHT
                );
                gvLigneEdit_CellValueChanged(view, args);
            }
        }

        private void txtCours_EditValueChanged(object sender, EventArgs e)
        {
            if (!decimal.TryParse(txtCours.Text, out decimal cours))
                return; // On ne fait rien si la valeur est invalide

            for (int i = 0; i < gvLigneEdit.RowCount; i++)
            {
                // Ignorer les lignes invalides (DevExpress utilise parfois des index spéciaux)
                if (gvLigneEdit.IsGroupRow(i))
                    continue;

                decimal qte = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(i, "DL_Qte") ?? 0);
                decimal pu = Convert.ToDecimal(gvLigneEdit.GetRowCellValue(i, "DL_PrixUnitaire") ?? 0);

                decimal montant = qte * pu * cours;

                gvLigneEdit.SetRowCellValue(i, "DL_MontantHT", montant);
            }

            // Regénère le résumé sans appeler l’event manuellement
            gvLigneEdit.UpdateSummary();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {
                using (AppDbContext context = new AppDbContext())
                {
                    // Pour un AJOUT (vérifier que l'enregistrement n'existe pas déjà)
                    var existingFRET = context.F_FRETS
                        .FirstOrDefault(s => s.DO_PIECE == dopiecetxt.Text);

                    if (existingFRET == null)
                    {
                        try
                        {
                            F_FRET f = new F_FRET();
                            f.DO_PIECE=dopiecetxt.Text;
                            f.DO_MONTANT = Convert.ToDecimal(txt_prix.Text);
                            f.DO_POIDS = Convert.ToDecimal(txt_poids.Text);

                            context.F_FRETS.Add(f);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    else
                    {
                        existingFRET.DO_MONTANT = Convert.ToDecimal(txt_prix.Text);
                        existingFRET.DO_POIDS = Convert.ToDecimal(txt_poids.Text);
                    }

                    context.SaveChanges();
                    MessageBox.Show("Modification FRET terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Data.Entity.Infrastructure.DbUpdateException ex)
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur");
            }
        }

        private void txt_prix_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Autoriser touches de contrôle (Backspace, etc.)
            if (char.IsControl(e.KeyChar))
                return;

            // Si l'utilisateur tape un point '.', on le remplace par une virgule ','
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';      // Remplacement automatique
            }

            // Autoriser les chiffres
            if (char.IsDigit(e.KeyChar))
                return;

            // Autoriser une seule virgule
            if (e.KeyChar == ',' && !txt_prix.Text.Contains(","))
                return;

            // Sinon : blocage de la saisie
            e.Handled = true;
        }

        private void txt_poids_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Autoriser touches de contrôle (Backspace, etc.)
            if (char.IsControl(e.KeyChar))
                return;

            // Si l'utilisateur tape un point '.', on le remplace par une virgule ','
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';      // Remplacement automatique
            }

            // Autoriser les chiffres
            if (char.IsDigit(e.KeyChar))
                return;

            // Autoriser une seule virgule
            if (e.KeyChar == ',' && !txt_poids.Text.Contains(","))
                return;

            // Sinon : blocage de la saisie
            e.Handled = true;
        }

        private void txt_poids_EditValueChanged(object sender, EventArgs e)
        {
            gvLigneEdit.RefreshData();
        }

        private void txt_prix_EditValueChanged(object sender, EventArgs e)
        {
            gvLigneEdit.RefreshData();
        }

        private string recup_ctnum(string cond)
        {
            string val = "";
            var recup = _context.F_COMPTET
                               .FirstOrDefault(x => x.CT_Intitule == cond);

            if (recup != null)
            {
                val= recup.CT_Num;
            }
            return val;
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            int? DE_No = Convert.ToInt32(lkDepot.EditValue);
            for (int i = 0; i < gvLigneEdit.RowCount; i++)
            {
                string reference = gvLigneEdit.GetRowCellValue(i, "AR_Ref")?.ToString();
                var qte = gvLigneEdit.GetRowCellValue(i, "DL_Qte");
                string designation = gvLigneEdit.GetRowCellValue(i, "DL_Design")?.ToString();

                string connectionStringArbio = $"Server={serveripPrincipale};Database=ARBIOCHEM;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";

                using (SqlConnection connection = new SqlConnection(connectionStringArbio))
                {
                    connection.Open();

                    string sql = @"
                    SELECT TOP 1 AR_SuiviStock
                    FROM F_ARTICLE
                    WHERE AR_Ref IN (SELECT AR_Ref FROM F_ARTSTOCK WHERE AR_Ref= @AR_Ref AND DE_No = @DE_No)";

                    using (SqlCommand checkCmd = new SqlCommand(sql, connection))
                    {
                        checkCmd.Parameters.AddWithValue("@AR_Ref", reference);
                        checkCmd.Parameters.AddWithValue("@DE_No", DE_No);

                        object result = checkCmd.ExecuteScalar();

                        if (result != null && result != DBNull.Value)
                        {
                            int arSuiviStock = Convert.ToInt32(result);

                            if (arSuiviStock == 5)
                            {
                                frmLotSerie frmLotS = new frmLotSerie();
                                frmLotS.txtreference.Text = reference;
                                frmLotS.txtdepot.Text = lkDepot.Text.ToString();
                                frmLotS.txtdesignation.Text = designation;
                                frmLotS.txtqte.Text = qte.ToString();
                                frmLotS.textDE_NO.Text = lkDepot.EditValue.ToString();
                                frmLotS.txtligne.Text = recuperer_DLNo(dopiecetxt.Text, reference).ToString();
                                frmLotS.ShowDialog();
                            }
                        }
                    }
                }
            }
        }

        private int recuperer_DLNo(string doPiece, string arRef)
        {
            int dlNo = 0;

            string connectionStringArbio = $"Server={serveripPrincipale};Database=TRANSIT;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";

            using (SqlConnection connection = new SqlConnection(connectionStringArbio))
            {
                connection.Open();

                string sql = @"
                SELECT TOP 1 DL_No
                FROM F_DOCLIGNE
                WHERE AR_Ref = @AR_Ref
                AND DO_PIECE = @DO_PIECE";

                using (SqlCommand cmd = new SqlCommand(sql, connection))
                {
                    cmd.Parameters.Add("@AR_Ref", SqlDbType.VarChar).Value = arRef;
                    cmd.Parameters.Add("@DO_PIECE", SqlDbType.VarChar).Value = doPiece;

                    object result = cmd.ExecuteScalar();

                    if (result != null && result != DBNull.Value)
                    {
                        dlNo = Convert.ToInt32(result);
                    }
                }
            }

            return dlNo;
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";

            if (string.IsNullOrWhiteSpace(lbl_val.Text))
            {
                lbl_val.Text = "1";
                string insertSql = @"
                    INSERT INTO F_MODE_PAIEMENT
                    (
                        do_piece,
                        mode_paiement,
                        date_echeance
                    )
                    VALUES
                    (
                        @do_piece,
                        @mode_paiement,
                        @date_echeance
                    )";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                using (SqlCommand cmd = new SqlCommand(insertSql, conn))
                {
                    // Paramètres TYPÉS (bonne pratique)
                    cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                        .Value = dopiecetxt.Text.Trim();

                    cmd.Parameters.Add("@mode_paiement", SqlDbType.VarChar, 20)
                        .Value = cmbmdp.Text?.ToString();
                    // OU cmbmdp.Text selon ton binding

                    cmd.Parameters.Add("@date_echeance", SqlDbType.DateTime)
                        .Value = dtecheance.Value;

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            else
            {
                lbl_val.Text = "1";
                string UPDATESql = @"
                   UPDATE F_MODE_PAIEMENT
                        SET mode_paiement=@mode_paiement,
                        date_echeance=@date_echeance
                   WHERE do_piece=@do_piece 
                   ";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                using (SqlCommand cmd = new SqlCommand(UPDATESql, conn))
                {
                    // Paramètres TYPÉS (bonne pratique)
                    cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                        .Value = dopiecetxt.Text.Trim();

                    cmd.Parameters.Add("@mode_paiement", SqlDbType.VarChar, 20)
                        .Value = cmbmdp.Text?.ToString();
                    // OU cmbmdp.Text selon ton binding

                    cmd.Parameters.Add("@date_echeance", SqlDbType.DateTime)
                        .Value = dtecheance.Value;

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";

            if (string.IsNullOrWhiteSpace(lblval1.Text))
            {
                lblval1.Text = "1";
                string insertSql = @"
                    INSERT INTO F_PACKING_LIST
                    (
                        dopiece,
                        type,
                        nombre,
                        eta,
                        etd
                    )
                    VALUES
                    (
                        @do_piece,
                        @type,
                        @nombre,
                        @eta,
                        @etd
                    )";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                using (SqlCommand cmd = new SqlCommand(insertSql, conn))
                {
                    cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                        .Value = dopiecetxt.Text.Trim();

                    cmd.Parameters.Add("@type", SqlDbType.VarChar, 20)
                        .Value = cmb_type.Text?.ToString();

                    cmd.Parameters.Add("@nombre", SqlDbType.VarChar)
                        .Value = txtnbr.Text;

                    cmd.Parameters.Add("@eta", SqlDbType.DateTime)
                        .Value = dteta.Value;

                    cmd.Parameters.Add("@etd", SqlDbType.DateTime)
                        .Value = dtetd.Value;

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            else
            {
                lblval1.Text = "1";
                string UPDATESql = @"
                UPDATE F_PACKING_LIST
                    SET type=@type,
                    nombre=@nombre,
                    eta=@eta,
                    etd=@etd
                    WHERE dopiece=@do_piece 
                ";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                    using (SqlCommand cmd = new SqlCommand(UPDATESql, conn))
                    {
                        cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                            .Value = dopiecetxt.Text.Trim();

                        cmd.Parameters.Add("@type", SqlDbType.VarChar, 20)
                            .Value = cmb_type.Text?.ToString();

                        cmd.Parameters.Add("@nombre", SqlDbType.VarChar)
                            .Value = txtnbr.Text;

                        cmd.Parameters.Add("@eta", SqlDbType.DateTime)
                            .Value = dteta.Value;

                        cmd.Parameters.Add("@etd", SqlDbType.DateTime)
                            .Value = dtetd.Value;

                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
    }
}