using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Models;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmReglementFrns : DevExpress.XtraEditors.XtraForm
    {
        private AppDbContext _context;
        private  string _rgLibelle;
        private string _fullRgLibelle;
        private frmEditDocument _parentForm;
        private decimal TotalReglement;
        private BindingList<F_CREGLEMENT> _bindingRegle;
        public frmReglementFrns(string RG_Libelle, decimal totalReglement, frmEditDocument parentForm)
        {
            InitializeComponent();
            _context = new AppDbContext();
            _rgLibelle = RG_Libelle;
            _fullRgLibelle = RG_Libelle;
            _context.Configuration.ProxyCreationEnabled = false;
            _parentForm = parentForm;
            TotalReglement = totalReglement;
        }

        private void frmReglementFrns_Load(object sender, EventArgs e)
        {
            try
            {
                _context = new AppDbContext();
                _context.Configuration.ProxyCreationEnabled = false;
                string Piece = _rgLibelle.Split('_').First();

                var deuxiemePartie = _rgLibelle.Split('_')
                                          .Skip(1) // saute la première partie
                                          .FirstOrDefault();

                var troisiemePartie = _rgLibelle.Split('_')
                                          .Skip(2) // saute la première partie
                                          .FirstOrDefault();





                Piece = Piece.Substring(Piece.Length - 8) + "_" + deuxiemePartie + "_" + troisiemePartie;
                var reglementList = _context.F_CREGLEMENT
                            .Where(f => f.RG_Libelle.Contains(Piece))
                            .ToList();
                gvRegle.Columns.Clear();

                _bindingRegle = new BindingList<F_CREGLEMENT>(reglementList);
                gcRegle.DataSource = _bindingRegle;
                gvRegle.OptionsBehavior.EditingMode = DevExpress.XtraGrid.Views.Grid.GridEditingMode.Inplace;
                CacherColonnes();

                gvRegle.Columns["RG_Montant"].Summary.Clear(); 
                gvRegle.Columns["RG_Montant"].Summary.Add(
                    DevExpress.Data.SummaryItemType.Sum,
                    "RG_Montant",
                    "Total = {0:n2}" 
                );

                gvRegle.Columns["RG_Montant"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                gvRegle.Columns["RG_Montant"].DisplayFormat.FormatString = "n2";
                gvRegle.Columns["RG_Cours"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                gvRegle.Columns["RG_Cours"].DisplayFormat.FormatString = "n2";


                //---------------------------------------------------------------------------------REPOSITORY LOOKUPEDIT
                
                
                var listReglement = _context.P_REGLEMENT.ToList();
                var listDevise = _context.P_DEVISE.ToList();
                var listCompteG = _context.F_COMPTEG.ToList();
                var listJournal = _context.F_JOURNAUX.ToList();

                var lookupReglement = new RepositoryItemLookUpEdit
                {
                    DataSource = listReglement,
                    DisplayMember = "R_Intitule",
                    ValueMember = "cbIndice",
                    NullText = "[Sélectionner...]"
                };
                // ⚠️ IMPORTANT : Vider les colonnes auto-générées
                lookupReglement.Columns.Clear();
                //lookupReglement.Columns.Add(new LookUpColumnInfo("N_Reglement", 0, "Code"));
                lookupReglement.Columns.Add(new LookUpColumnInfo("R_Intitule", 0, "Mode de règlement"));
                gvRegle.GridControl.RepositoryItems.Add(lookupReglement);
                gvRegle.Columns["N_Reglement"].ColumnEdit = lookupReglement;

                var lookupDevise = new RepositoryItemLookUpEdit
                {
                    DataSource = listDevise,
                    DisplayMember = "D_Intitule",
                    ValueMember = "cbIndice",
                    NullText = "[Sélectionner...]"
                };
                lookupDevise.Columns.Clear();
                //lookupDevise.Columns.Add(new LookUpColumnInfo("N_Devise", 0, "Code"));
                lookupDevise.Columns.Add(new LookUpColumnInfo("D_Intitule", 0, "Devise"));
                gvRegle.GridControl.RepositoryItems.Add(lookupDevise);
                gvRegle.Columns["N_Devise"].ColumnEdit = lookupDevise;

                var lookupCompteG = new RepositoryItemLookUpEdit
                {
                    DataSource = listCompteG,
                    DisplayMember = "CG_Num",
                    ValueMember = "CG_Num",
                    NullText = "[Sélectionner...]"
                };
                lookupCompteG.Columns.Clear();
                lookupCompteG.Columns.Add(new LookUpColumnInfo("CG_Num", 0, "Code"));
                gvRegle.GridControl.RepositoryItems.Add(lookupCompteG);
                gvRegle.Columns["CG_Num"].ColumnEdit = lookupCompteG;

                var lookupJournal = new RepositoryItemLookUpEdit
                {
                    DataSource = listJournal,
                    DisplayMember = "JO_Intitule",
                    ValueMember = "JO_Num",
                    NullText = "[Sélectionner...]"
                };
                lookupJournal.Columns.Clear();
                lookupJournal.Columns.Add(new LookUpColumnInfo("JO_Num", 0, "Journal"));
                lookupJournal.Columns.Add(new LookUpColumnInfo("JO_Intitule", 0, "Intitulé"));
                gvRegle.GridControl.RepositoryItems.Add(lookupJournal);
                gvRegle.Columns["JO_Num"].ColumnEdit = lookupJournal;







                gvRegle.BestFitColumns();   
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnNewRegle_Click(object sender, EventArgs e)
        {
            try
            {
                int count = _context.F_CREGLEMENT.Max(u => u.RG_No).Value;

                string CT_NumPayeur = _rgLibelle.Split('_').Last();

                string Piece = _rgLibelle.Split('_').First();

                //dlpiece = dlpiece.Substring(dlpiece.Length - 8);
                //string rglibelle = dlpiece + "_" + arref + "_" + ctnum;



                // Création d'un nouveau règlement avec initialisation minimale
                var newRegle = new F_CREGLEMENT
                {
                    RG_No = count + 1,
                    CT_NumPayeur = CT_NumPayeur,
                    RG_Date = DateTime.Now,
                    RG_Reference = "",
                    RG_Libelle = _fullRgLibelle ,
                    RG_Montant = 0,
                    RG_MontantDev = 0,
                    N_Reglement = 1,
                    RG_Impute = 0,
                    RG_Compta = 0,
                    EC_No = 0,
                    RG_Type = 1,
                    RG_Cours = 0,
                    N_Devise = 1,
                    JO_Num = "ACHT",
                    CG_Num = "40120000",
                    RG_TypeReg = 0,
                    RG_Heure = "000" + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second,
                    RG_Piece = (count + 1).ToString("00000"),
                    CA_No = null,
                    cbCA_No = null,
                    CO_NoCaissier = frmEditDocument.intcollaborateur,  //----------------------------------------------ACHETEUR IHANY
                    cbCO_NoCaissier = frmEditDocument.intcollaborateur,
                    RG_Banque = 0,
                    RG_Transfere = 0,
                    RG_Cloture = 0,
                    RG_Ticket = 0,
                    RG_Souche = 0,
                    CT_NumPayeurOrig = CT_NumPayeur,
                    RG_DateEchCont = new DateTime(1753, 1, 1, 0, 0, 0, 0),
                    CG_NumEcart =null,
                    JO_NumEcart = null,
                    RG_MontantEcart = 0,
                    RG_NoBonAchat=0,
                    RG_Valide = 0,
                    RG_Anterieur = 0,
                    RG_MontantCommission = 0,
                    RG_MontantNet = 0,
                    cbModification = DateTime.Now,
                    cbCreateur = "COLS", // si tu veux enregistrer l’utilisateur courant
                    cbCreation = DateTime.Now,
                    cbReplication = 0,
                    cbFlag = 0,
                    cbHashVersion = 1,
                    cbHash = null,
                    cbHashDate = null,
                    cbHashOrder = null,
                    cbProt = 0
                };

                // Ajout à la liste liée à l'UI
                _bindingRegle.Add(newRegle);

                // Ajout au contexte EF pour le suivi
                _context.F_CREGLEMENT.Add(newRegle);

                // Déplace le curseur sur la nouvelle ligne et entre en mode édition
                var newRowHandle = gvRegle.RowCount - 1;
                gvRegle.FocusedRowHandle = newRowHandle;
                gvRegle.FocusedColumn = gvRegle.VisibleColumns[0];
                gvRegle.ShowEditor();

                // Rendre certaines colonnes en lecture seule si besoin
                gvRegle.Columns["cbModification"].OptionsColumn.ReadOnly = true;
                gvRegle.Columns["cbCreateur"].OptionsColumn.ReadOnly = true;

                


                gvRegle.BestFitColumns(); // Ajustement automatique des colonnes
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void CacherColonnes()
        {
            gvRegle.Columns["cbModification"].VisibleIndex = -1;
            gvRegle.Columns["cbCreateur"].VisibleIndex = -1;
            gvRegle.Columns["cbCreation"].VisibleIndex = -1;
            gvRegle.Columns["cbReplication"].VisibleIndex = -1;
            gvRegle.Columns["cbFlag"].VisibleIndex = -1;
            gvRegle.Columns["cbHashVersion"].VisibleIndex = -1;
            gvRegle.Columns["cbHash"].VisibleIndex = -1;
            gvRegle.Columns["cbHashDate"].VisibleIndex = -1;
            gvRegle.Columns["cbHashOrder"].VisibleIndex = -1;
            gvRegle.Columns["cbProt"].VisibleIndex = -1;
            gvRegle.Columns["cbCA_No"].VisibleIndex = -1;
            gvRegle.Columns["cbCO_NoCaissier"].VisibleIndex = -1;
            gvRegle.Columns["RG_Cloture"].VisibleIndex = -1;
            gvRegle.Columns["RG_Compta"].VisibleIndex = -1;
            gvRegle.Columns["RG_Transfere"].VisibleIndex = -1;
            gvRegle.Columns["RG_Ticket"].VisibleIndex = -1;
            gvRegle.Columns["RG_Souche"].VisibleIndex = -1;
            gvRegle.Columns["RG_NoBonAchat"].VisibleIndex = -1;
            gvRegle.Columns["RG_Anterieur"].VisibleIndex = -1;
            gvRegle.Columns["RG_MontantCommission"].VisibleIndex = -1;
            gvRegle.Columns["RG_MontantNet"].VisibleIndex = -1;
            gvRegle.Columns["CT_NumPayeurOrig"].VisibleIndex = -1;
            gvRegle.Columns["CG_NumEcart"].VisibleIndex = -1;
            gvRegle.Columns["JO_NumEcart"].VisibleIndex = -1;
            gvRegle.Columns["RG_DateEchCont"].VisibleIndex = -1;
            //gvRegle.Columns["N_Reglement"].VisibleIndex = -1;
            gvRegle.Columns["EC_No"].VisibleIndex = -1;
            //gvRegle.Columns["N_Devise"].VisibleIndex = -1;
            //gvRegle.Columns["RG_Cours"].VisibleIndex = -1;
            gvRegle.Columns["CA_No"].VisibleIndex = -1;
            gvRegle.Columns["CO_NoCaissier"].VisibleIndex = -1;
            gvRegle.Columns["RG_Banque"].VisibleIndex = -1;
            gvRegle.Columns["cbMarq"].VisibleIndex = -1;
            gvRegle.Columns["RG_Piece"].VisibleIndex = -1;
            //gvRegle.Columns["JO_Num"].VisibleIndex = -1;
            //gvRegle.Columns["CG_Num"].VisibleIndex = -1;
            gvRegle.Columns["RG_TypeReg"].VisibleIndex = -1;
            gvRegle.Columns["RG_Type"].VisibleIndex = -1;
            gvRegle.Columns["cbRG_Piece"].VisibleIndex = -1;
            //gvRegle.Columns["RG_Libelle"].VisibleIndex = -1;
            gvRegle.Columns["RG_Impute"].VisibleIndex = -1;
            gvRegle.Columns["RG_Heure"].VisibleIndex = -1;
            //gvRegle.Columns["RG_Reference"].VisibleIndex = -1;
            gvRegle.Columns["RG_MontantDev"].VisibleIndex = -1;
            //gvRegle.Columns["RG_Montant"].VisibleIndex = -1;
            //gvRegle.Columns["RG_Date"].VisibleIndex = -1;
            gvRegle.Columns["cbCA_No"].VisibleIndex = -1;
            gvRegle.Columns["cbCO_NoCaissier"].VisibleIndex = -1;
            gvRegle.Columns["cbProt"].VisibleIndex = -1;
            gvRegle.Columns["RG_No"].VisibleIndex = -1;
            gvRegle.Columns["cbCT_NumPayeur"].VisibleIndex = -1;
            gvRegle.Columns["cbEC_No"].VisibleIndex = -1;
            gvRegle.Columns["CG_NumCont"].VisibleIndex = -1;
            gvRegle.Columns["cbCG_NumCont"].VisibleIndex = -1;
            gvRegle.Columns["RG_Impaye"].VisibleIndex = -1;
            gvRegle.Columns["cbCG_Num"].VisibleIndex = -1;
            gvRegle.Columns["RG_MontantEcart"].VisibleIndex = -1;
            gvRegle.Columns["RG_Valide"].VisibleIndex = -1;
            gvRegle.Columns["cbCreationUser"].VisibleIndex = -1;
            gvRegle.Columns["cbCT_NumPayeurOrig"].VisibleIndex = -1;
            gvRegle.Columns["cbCG_NumEcart"].VisibleIndex = -1;
        }

        private void btnSaveRegle_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (var reglement in _bindingRegle)
                {
                    // Mise à jour de la date de modification
                    reglement.cbModification = DateTime.Now;
                    reglement.cbCreationUser = FrmMdiParent._id_user;

                    // Vérifie s'il s'agit d'un nouveau règlement (cbMarq = 0 = pas encore en DB)
                    if (reglement.cbMarq == 0)
                    {
                        _context.F_CREGLEMENT.Add(reglement);
                    }
                    else
                    {
                        // L'entité vient du contexte, elle est déjà trackée : aucune action nécessaire
                        // Si elle a été modifiée, EF détectera automatiquement les changements
                    }
                }

                _context.SaveChanges();

                _parentForm.UPdateLigne(_parentForm.laligneamettreajour);

                MessageBox.Show("Les règlements ont été enregistrés avec succès.",
                                "Enregistrement",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}