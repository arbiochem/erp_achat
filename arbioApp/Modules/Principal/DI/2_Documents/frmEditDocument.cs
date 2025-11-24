using DevExpress.XtraEditors;
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
using arbioApp.Modules.Principal.Dashboard.CIAL;
using DevExpress.XtraGrid;
using arbioApp.Models;
using System.Reflection;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraGrid.Views.Grid;
using System.Globalization;
using DevExpress.XtraEditors.Controls;
using Objets100cLib;
using arbioApp.DTO;
using arbioApp.Services;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Models.Json;
using arbioApp.Modules.Principal.DI.Services;
using DevExpress.DataAccess.DataFederation;
using DevExpress.ChartRangeControlClient.Core;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using DevExpress.XtraBars.Customization;
using DevExpress.Data;
using DevExpress.XtraReports.UI;
////using Microsoft.Office.Interop.Outlook;
using System.Net;
using DevExpress.Pdf.Xmp;
using System.Data.Entity;
using DevExpress.DataProcessing.InMemoryDataProcessor.GraphGenerator;
using DevExpress.XtraExport.Helpers;
using DevExpress.DashboardWin.Design;
using Exception = System.Exception;
using DevExpress.XtraGrid.Columns;
using DevExpress.DataAccess.Excel;
using System.IO;
using DevExpress.Spreadsheet;
using FieldInfo = DevExpress.DataAccess.Excel.FieldInfo;
using DevExpress.DataAccess.UI.Excel;
using System.Collections;
using DevExpress.XtraSplashScreen;
using System.Threading;
using DevExpress.DataAccess.Sql;
using DevExpress.CodeParser;
using DevExpress.Utils.About;
using DevExpress.XtraTab;
using arbioApp.Modules.Helpers;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraSpreadsheet;
using System.Security.AccessControl;
using DevExpress.XtraPrinting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DevExpress.XtraSpreadsheet.Import.Xls;
using Syncfusion.Windows.Forms.Maps;


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
            _listeDevise = Entetes.GetAllDevise();

            lkDevise.Properties.DataSource = _listeDevise;
            lkDevise.Properties.ValueMember = "cbMarq"; // Clé réelle stockée
            lkDevise.Properties.DisplayMember = "D_Intitule"; // Texte affiché
            lkDevise.Properties.PopulateColumns();
            lkDevise.Properties.Columns.Clear();

            lkDevise.Properties.Columns.Add(new LookUpColumnInfo("cbMarq", "cbMarq", 50));
            lkDevise.Properties.Columns.Add(new LookUpColumnInfo("D_Intitule", "DEVISE"));

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

            if (gvLigneEdit.Columns["DL_Frais"] != null)
            {
                gvLigneEdit.Columns["DL_Frais"].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Custom;
                gvLigneEdit.Columns["DL_Frais"].SummaryItem.DisplayFormat = "{0:n2}";
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


                    DataTable dt = (DataTable)gcLigneEdit.DataSource;

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
                    newRow["DL_Qte"] = Qte_propice;
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




                    _f_DOCLIGNEService.UpdateF_DOCLIGNE(dopiecetxt.Text, CtNum, arRef, dl_Designe, puBrut, dlLigne, quantiteEcriteStock, _typeDocument, dlTaxe1, dlMontantHT, dlMontantTTC, retenu, remisePourcent, dlPiecefrns, dlDatePiecefrns, montantRegl);

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
            UPdateLigne(row);
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
        private void Hyperlink_Click_AddLigne(object sender, EventArgs e)
        {
            AddLigne();
        }

        public void ExecuteStockAlert()
        {
            string query1 = @"SELECT * FROM dbo.VW_ETAT_STOCK";

            string connectionStringArbapp = $"Server={FrmMdiParent.DataSourceNameValueParent};" +
                                            $"Database=arbapp;User ID=Dev;Password=1234;" +
                                            $"TrustServerCertificate=True;Connection Timeout=120;";

            DataTable dtMaster =
                arbioApp.Modules.Principal.BrowsSites.ExecuteQueryOnMultipleServers(connectionStringArbapp, query1);

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

            foreach (var group in groupedData)
            {
                // Calcul des sommes des enfants
                decimal totalStockReel = group.Sum(r => r.Field<decimal>("STOCK REEL"));
                decimal totalStockMini = group.Sum(r => r.Field<decimal>("STOCK MINI"));
                decimal totalStockMaxi = group.Sum(r => r.Field<decimal>("STOCK MAXI"));

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



            treeList1.EndUpdate();
        }
        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            frmSites frmsite = new frmSites(this);
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

        private void btnValider_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValiderChampsObligatoires())
                    return;

                F_DOCENTETE doc = _f_DOCENTETEService.GetDocByPiece(dopiecetxt.Text, _listeDocs);
                string _currentDocPieceNo = dopiecetxt.Text;
                if (doc != null)
                {
                    UpdateFDOCENTETE();
                }
                else
                {
                    InsertFDOCENTETE();
                }

                _ucDocuments.RafraichirDonnees();
                gvLigneEdit.SetFocusedValue(lkEdFrns.EditValue);
                StatutActuel = Convert.ToInt32(lkStatut.EditValue);
            }

            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
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

                int number = int.Parse(dopiecetxt.Text.Substring(dopiecetxt.Text.Length - 4));
                decimal DoTaxe1 = string.IsNullOrEmpty(doTaxe1txt.Text)? 0: Convert.ToDecimal(doTaxe1txt.Text);

                short DoStatut = Convert.ToInt16(lkStatut.EditValue);

                if (fDocenteteToModif == null)
                {
                    _f_DOCENTETEService.UpdateProprietesF_DOCENTETE(_currentDocPieceNo,
    dateLivrPrevu, dateLivrReal, reference, caNum, CO_NO,
    (short?)numExpedition, expeditIntitule, dO_Coord01, numeroDepot, DoTaxe1, DoStatut, ctnum);
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
            try
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

                if (fDocenteteToModif == null)
                {
                    _f_DOCENTETEService.InsertNewF_DOCENTETE(_typeDocument, _currentDocPieceNo, frns,
                        numExpedition, caNum, caisseNumber,
                        caissierNumber, expeditIntitule, dateLivrPrevu,
                        dateLivrReal, reference, dO_Coord01,
                        (int)representant.CO_No, numeroDepot, deviseID,
                        ct_taux2, CodeTaxe, _prefix,
                        number, DoTaxe1);
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
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



        private void gvLigneEdit_ShownEditor(object sender, EventArgs e)
        {

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


        private void gvLigneEdit_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {

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
                || e.Column.FieldName == "DL_Taxe1" || e.Column.FieldName == "DL_PUDevise")
            {
                GridView view = sender as GridView;

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

                    // Appliquer la remise au prix unitaire
                    decimal prixNet = prixUnitaire * (1 - remise / 100);
                    decimal montantHT = prixNet * quantite;
                    decimal montantTTC = montantHT * (1 + tauxTVA / 100);

                    view.SetRowCellValue(e.RowHandle, "DL_MontantHT", montantHT);
                    view.SetRowCellValue(e.RowHandle, "DL_MontantTTC", montantTTC);

                }
            }


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
            if (view.FocusedColumn.FieldName == "DL_Qte")
            {
                UPdateLigne(row);

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


        public static int intcollaborateur;
        private void frmEditDocument_Load_1(object sender, EventArgs e)
        {
            gvLigneEdit.ShowingEditor -= gvLigneEdit_ShowingEditor; // Se désabonner au cas où
            intcollaborateur = cono;
            ShapeFileLayer shapeLayer = new ShapeFileLayer();

            shapeLayer.Uri = "world1.shp";

            this.maps1.Layers.Add(shapeLayer);
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
            if (StatutActuel < 2)
            {
                MessageBox.Show("Le statut actuel ne permet pas de transformer le document.", "Information");
                return;
            }
            var dlg = new frmTransform(_typeDocument);
            dlg.ParentFormInstance = this;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                this.TransformFDOCENTETE(dlg.doctype);
            }

            // Fix for CS0120: Use the instance of ucDocuments instead of trying to call it statically
            _ucDocuments.ChargerDonneesDepuisBDD();
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
            MettreAJourFiltre();
        }

        private void chkSommeil_CheckedChanged(object sender, EventArgs e)
        {
            MettreAJourFiltre();
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


                frmEditInfos frm = new frmEditInfos(dopiece, totalTVA, this);
                frm.ShowDialog();
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
                xtraOpenFileDialog1.Filter = "Microsoft Excel Files (*.xlsx)|*.xlsx";
                if (xtraOpenFileDialog1.ShowDialog() != DialogResult.OK) return;

                string filePath = xtraOpenFileDialog1.FileName;

                // Charger les données depuis Excel
                Workbook workbook = new Workbook();
                workbook.LoadDocument(filePath);
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


                if (dataDocLigneImport.Rows.Count == 0)
                {
                    XtraMessageBox.Show("Aucune ligne trouvée dans le fichier Excel.");
                    return;
                }

                SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    int total = dataDocLigneImport.Rows.Count;
                    int current = 0;

                    foreach (DataRow row in dataDocLigneImport.Rows)
                    {
                        current++;
                        SplashScreenManager.Default.SetWaitFormDescription($"{current}/{total}");


                        int dotype = 10;
                        string arRef = row["AR_Ref"]?.ToString() ?? "";
                        string ctNum = row["CT_Num"]?.ToString() ?? "";

                        decimal dlQte = row["DL_Qte"] != DBNull.Value ? Convert.ToDecimal(row["DL_Qte"]) : 0;

                        decimal DLTaxe1 = 0;
                        F_ARTICLE articleChoisi = _f_ARTICLERepository.GetF_ARTICLEByAR_Ref(arRef);
                        int? DE_No = Convert.ToInt32(lkDepot.EditValue);
                        int? dl_Ligne = 0;
                        decimal previousQuantite = 0;
                        decimal remisePourcent = 0;
                        string arDesign = row["DL_Design"]?.ToString() ?? "";
                        decimal puBrut = row["DL_PrixUnitaire"] != DBNull.Value ? Convert.ToDecimal(row["DL_PrixUnitaire"]) : 0;
                        decimal puNet = puBrut * (1 - remisePourcent / 100);
                        decimal montantHT = dlQte * puNet;
                        decimal montantTTC = montantHT * (1 + DLTaxe1 / 100);
                        int retenu = 1;
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
                        //string doAffaire = caNum;

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


                        //UPDATE LES PRIX DANS F_DOCENTETE



                    }
                    MettreAJourTotauxDepuisBD(dopiecetxt.Text);
                    conn.Close();
                    string _currentDocPieceNo = dopiecetxt.Text;
                    InitializeGrid(gcLigneEdit, _currentDocPieceNo);
                }

                SplashScreenManager.CloseForm();
                XtraMessageBox.Show("Importation terminée avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                SplashScreenManager.CloseForm();
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);            
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
            if (e.NewValue == null) return;
            int newValue = Convert.ToInt32(e.NewValue);

            if (newValue < StatutActuel)
            {
                e.Cancel = true;
                MessageBox.Show("Impossible de revenir à un statut précédent.", "Action non autorisée", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            PointF latLong = this.maps1.GetMapCoordinates(point);

            float latitude = latLong.Y;
            float longitude = latLong.X;

            // Affiche dans la barre de statut ou un Label
            lblCoord.Text = $"Lat: {latitude:F4}, Lon: {longitude:F4}";

            // Ou dans un tooltip
            System.Windows.Forms.ToolTip tt = new System.Windows.Forms.ToolTip();
            tt.SetToolTip(maps1, $"Lat: {latitude:F4}, Lon: {longitude:F4}");
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
            FiltrerArticles();
        }

        private void chkArSommeil_CheckedChanged(object sender, EventArgs e)
        {
            FiltrerArticles();
        }

        private void hyperlinkLabelControl8_Click(object sender, EventArgs e)
        {
            ChargerArticles();
        }
    }
}