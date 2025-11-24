using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Models;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using arbioApp.Utils.Connection;
using DevExpress.CodeParser;
using DevExpress.Xpo.DB.Helpers;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraRichEdit.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Validation;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmEditInfos : DevExpress.XtraEditors.XtraForm
    {
        private readonly string _doPiece;
        private AppDbContext _context;
        private BindingList<F_DOCFRAISIMPORT> _bindingFrais;
        private frmEditDocument _parentForm;
        private decimal TotalTVA;
        public frmEditInfos(string doPiece, decimal totalTVA, frmEditDocument parentForm)
        {
            InitializeComponent();
            _doPiece = doPiece;
            _context = new AppDbContext();
            _context.Configuration.ProxyCreationEnabled = false;
            _parentForm = parentForm;
            TotalTVA = totalTVA;
        }
        public static decimal totalMontantFrais;
        private void frmEditInfos_Load(object sender, EventArgs e)
        {
            //Load FRET
           
            var fret = _context.F_FRETS
            .FirstOrDefault(f => f.DO_PIECE == _doPiece);

            if (fret != null)
            {
                // Affichage dans les TextBox
                //txt_prix.Text = fret.DO_PRIX.ToString("0.00");
                //txt_montant.Text = fret.DO_MONTANT.ToString("0.00");
                //txt_poids.Text = fret.DO_POIDS.ToString("0.00");
            }

            LoadData();
            GridView gridView = gridControlDevises.MainView as GridView;

            // Se positionner sur la nouvelle ligne et activer l'édition
            int newRowHandle = gridView.RowCount - 1;
            gridView.FocusedRowHandle = newRowHandle;
            gridView.FocusedColumn = gridView.VisibleColumns[0];
            gridView.ShowEditor();

            if (gridView != null)
            {
                // 1. Activer l’édition et la ligne d’ajout
                gridView.OptionsBehavior.Editable = true;
                gridView.OptionsView.NewItemRowPosition = NewItemRowPosition.Top;

                // 2. Fournir une source vide si nécessaire
                if (fDEVISEBindingSource.DataSource == null)
                {
                    fDEVISEBindingSource.DataSource = new List<F_DEVISE>();
                    gridControlDevises.DataSource = fDEVISEBindingSource;
                }

                // 3. Forcer la création de la ligne d’ajout et activer l’éditeur
                gridView.FocusedRowHandle = GridControl.NewItemRowHandle;
                gridView.FocusedColumn = gridView.VisibleColumns[0]; // première colonne éditable
                gridView.ShowEditor(); // ouvre l’éditeur de cellule
                gridView.UpdateCurrentRow();

            }

            // TODO: cette ligne de code charge les données dans la table 'aRBIOCHEMDataSet.F_DEVISE'. Vous pouvez la déplacer ou la supprimer selon les besoins.
            this.f_DEVISETableAdapter.Fill(this.aRBIOCHEMDataSet.F_DEVISE);
            try
            {
                _context = new AppDbContext();

                _context.Configuration.ProxyCreationEnabled = false;

                var fraisList = _context.F_DOCFRAISIMPORT
                        .Include("TypeFrais")
                        .Include("Repartition")
                        .Where(f => f.DO_Piece == _doPiece)
                        .ToList();

                gridViewFrais.Columns.Clear();

                _bindingFrais = new BindingList<F_DOCFRAISIMPORT>(fraisList);
                gridControlFrais.DataSource = _bindingFrais;

                RepositoryItemMemoEdit memoEdit = new RepositoryItemMemoEdit();
                gridViewFrais.Columns["FI_Observation"].ColumnEdit = memoEdit;
                gridViewFrais.OptionsView.RowAutoHeight = true;


                /*if(txt_montant.Text != "")
                {
                    gridViewFrais.Columns["FI_Montant_AR"].Summary.Clear();
                    gridViewFrais.Columns["FI_Montant_AR"].Summary.Add(
                        DevExpress.Data.SummaryItemType.Sum,
                        "FI_Montant_AR",
                        "" 
                    );
                }
                else
                {
                    gridViewFrais.Columns["FI_Montant_AR"].Summary.Clear(); // Par sécurité
                    gridViewFrais.Columns["FI_Montant_AR"].Summary.Add(
                        DevExpress.Data.SummaryItemType.Sum,
                        "FI_Montant_AR",
                        "Total = {0:n2}" // Format exemple : 2 500.00
                    );
                }*/

                var totalMontant = _bindingFrais.Sum(f => f.FI_Montant);
                totalMontantFrais = totalMontant; // Stocke le total pour utilisation ultérieure

                gridViewFrais.Columns["FI_Montant"].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                gridViewFrais.Columns["FI_Montant"].DisplayFormat.FormatString = "n2";

                // Masquer les colonnes navigationnelles (évite les proxy affichés)
                gridViewFrais.Columns["TypeFrais"].Visible = false;
                gridViewFrais.Columns["Repartition"].Visible = false;


                if (gridViewFrais.DataRowCount > 0){
                    //txt_poids.Enabled = true;
                    //txt_prix.Enabled = true;
                }else{
                    //txt_poids.Enabled = false;
                    //txt_prix.Enabled = false;
                }
                //-----------------------------------------------------------------------------------------------

                var typeFraisList = _context.P_TYPEFRAIS.Where(x=>x.FR_Intitule != "FRET").ToList();

                // 2. Créer l'éditeur TreeListLookUpEdit
                var treeListLookup = new RepositoryItemTreeListLookUpEdit();

                treeListLookup.TreeList.BeforeFocusNode += (s, e) =>
                {
                    var tree = s as DevExpress.XtraTreeList.TreeList;
                    if (e.Node.HasChildren)
                    {
                        e.CanFocus = false;
                    }
                };


                treeListLookup.DisplayMember = "FR_Intitule";
                treeListLookup.ValueMember = "FR_Num";
                treeListLookup.NullText = "";

                // 3. Configurer la TreeList interne
                treeListLookup.TreeList.KeyFieldName = "FR_Num";
                treeListLookup.TreeList.ParentFieldName = "FR_ParentId";
                treeListLookup.DataSource = typeFraisList;
                

                // (Optionnel) ajuster les colonnes visibles
                treeListLookup.TreeList.Columns["Parent"].Visible = false;
                treeListLookup.TreeList.Columns["Enfants"].Visible = false;
                treeListLookup.TreeList.Columns["FraisImports"].Visible = false;




                gridViewFrais.Columns["FI_TypeFraisId"].ColumnEdit = treeListLookup;





                //-----------------------------------------------------------------------------------------------


                var repartitionList = _context.P_TYPEREPARTITION.ToList();
                var repartitionEditor = new RepositoryItemLookUpEdit
                {
                    DataSource = repartitionList,
                    DisplayMember = "RP_Intitule",
                    ValueMember = "RP_Num",
                    NullText = ""
                };
                repartitionEditor.PopulateColumns();
                repartitionEditor.Columns["FraisImports"].Visible = false;
                gridControlFrais.RepositoryItems.Add(repartitionEditor);
                gridViewFrais.Columns["FI_RepartitionId"].ColumnEdit = repartitionEditor;
                gridViewFrais.Columns["FI_Montant_AR"].OptionsColumn.AllowEdit = false;
                gridViewFrais.Columns["FI_Montant_AR"].OptionsColumn.ReadOnly = true;

                var deviseList = _context.P_DEVISE
                                            .Where(d => d.D_Intitule != "" && d.D_Intitule != null)
                                            .ToList();

                var deviseEditor = new RepositoryItemLookUpEdit
                {
                    DataSource = deviseList,
                    DisplayMember = "D_Intitule",   // Ce que l'utilisateur voit
                    ValueMember = "D_Intitule",     // Ce qui est stocké dans FI_Devise
                    NullText = ""
                };

                deviseEditor.PopulateColumns();
                deviseEditor.Columns.Clear();
                deviseEditor.Columns.Add(new LookUpColumnInfo("D_Intitule", 0, "Devise"));
                gridControlFrais.RepositoryItems.Add(deviseEditor);

                // Affecte l'éditeur à la colonne
                gridViewFrais.Columns["FI_Devise"].ColumnEdit = deviseEditor;
                gridViewFrais.BestFitColumns();
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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
        private void btnEnregistrer_Click(object sender, EventArgs e)
        {
            if (tester_cloturer(_doPiece))
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
                GridView view = gridViewFrais;

                Boolean pasVide = true;
                for (int i = 0; i < view.RowCount; i++)
                {
                    var value = view.GetRowCellValue(i, "FI_Piece");

                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                    {
                        MessageBox.Show($"Valeur vide pour FI_Pièce trouvée dans la ligne N° {i + 1}","Message d'erreur",MessageBoxButtons.OK,MessageBoxIcon.Error);
                        pasVide = false;
                        break;
                    }
                }

                if (pasVide)
                {
                    if (_doPiece.ToString().StartsWith("APA"))
                    {
                        bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");

                        if (autorise)
                        {
                            //Enregistrer Fret

                            // Recherche si le DO_PIECE existe déjà
                            var existingFret = _context.F_FRETS
                                .FirstOrDefault(f => f.DO_PIECE == _doPiece);

                            if (existingFret != null)
                            {
                                //existingFret.DO_PRIX = decimal.Parse(txt_prix.Text);
                                //existingFret.DO_MONTANT = decimal.Parse(txt_montant.Text);
                                //existingFret.DO_POIDS = decimal.Parse(txt_poids.Text);
                            }
                            else
                            {
                                var f_fret = new F_FRET
                                {
                                    DO_PIECE = _doPiece,
                                    //DO_PRIX = decimal.Parse(txt_prix.Text),
                                    //DO_MONTANT = decimal.Parse(txt_montant.Text),
                                    //DO_POIDS = decimal.Parse(txt_poids.Text)
                                };

                                _context.F_FRETS.Add(f_fret);
                            }

                            _context.SaveChanges();


                            try
                            {
                                gridViewFrais.CloseEditor();
                                gridViewFrais.UpdateCurrentRow();

                                _context.SaveChanges();

                                // Exécuter la procédure
                                string sql = "EXEC SP_CalculCoutRevientParValeur @DO_Piece, @prix_total,@poids_total";

                                _context.Database.ExecuteSqlCommand(
                                   sql,
                                   new SqlParameter("@DO_Piece", _doPiece),
                                   new SqlParameter("@prix_total", Convert.ToDecimal(_parentForm.txt_prix.Text)),
                                   new SqlParameter("@poids_total", Convert.ToDecimal(_parentForm.txt_poids.Text))
                               );


                                // Rafraîchir la grille du parent
                                _parentForm?.InitializeGrid(_parentForm.GridLigneEdit, _doPiece);


                                MessageBox.Show("Frais enregistrés et coût de revient mis à jour avec succès.");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Erreur : " + ex.Message);
                            }


                        }
                        else
                        {
                            MessageBox.Show(
                                "Vous n'avez pas l'autorisation de modifier un projet d'achat !",
                                "Modification bloquée",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                    else if (_doPiece.ToString().StartsWith("ABR"))
                    {
                        bool autorise = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                        if (autorise)
                        {
                            //Enregistrer Fret

                            // Recherche si le DO_PIECE existe déjà
                            var existingFret = _context.F_FRETS
                                .FirstOrDefault(f => f.DO_PIECE == _doPiece);

                            if (existingFret != null)
                            {
                                //existingFret.DO_PRIX = decimal.Parse(txt_prix.Text);
                                //existingFret.DO_MONTANT = decimal.Parse(txt_montant.Text);
                                //existingFret.DO_POIDS = decimal.Parse(txt_poids.Text);
                            }
                            else
                            {
                                var f_fret = new F_FRET
                                {
                                    DO_PIECE = _doPiece,
                                    //DO_PRIX = decimal.Parse(txt_prix.Text),
                                    //DO_MONTANT = decimal.Parse(txt_montant.Text),
                                    //DO_POIDS = decimal.Parse(txt_poids.Text)
                                };

                                _context.F_FRETS.Add(f_fret);
                            }

                            _context.SaveChanges();


                            try
                            {
                                gridViewFrais.CloseEditor();
                                gridViewFrais.UpdateCurrentRow();

                                _context.SaveChanges();

                                // Exécuter la procédure
                                string sql = "EXEC SP_CalculCoutRevientParValeur @DO_Piece, @prix_total,@poids_total";

                                _context.Database.ExecuteSqlCommand(
                                   sql,
                                   new SqlParameter("@DO_Piece", _doPiece),
                                   new SqlParameter("@prix_total", Convert.ToDecimal(_parentForm.txt_prix.Text)),
                                   new SqlParameter("@poids_total", Convert.ToDecimal(_parentForm.txt_poids.Text))
                               );

                                // Rafraîchir la grille du parent
                                _parentForm?.InitializeGrid(_parentForm.GridLigneEdit, _doPiece);


                                MessageBox.Show("Frais enregistrés et coût de revient mis à jour avec succès.");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Erreur : " + ex.Message);
                            }


                        }
                        else
                        {
                            MessageBox.Show(
                                "Vous n'avez pas l'autorisation de modifier un bon de réception !",
                                "Modification bloquée",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                    else if (_doPiece.ToString().StartsWith("AFA"))
                    {
                        bool autorise = frmMenuAchat.verifier_droit("Facture", "UPDATE");

                        if (autorise)
                        {
                            //Enregistrer Fret

                            // Recherche si le DO_PIECE existe déjà
                            var existingFret = _context.F_FRETS
                                .FirstOrDefault(f => f.DO_PIECE == _doPiece);

                            if (existingFret != null)
                            {
                                //existingFret.DO_PRIX = decimal.Parse(txt_prix.Text);
                                //existingFret.DO_MONTANT = decimal.Parse(txt_montant.Text);
                                //existingFret.DO_POIDS = decimal.Parse(txt_poids.Text);
                            }
                            else
                            {
                                var f_fret = new F_FRET
                                {
                                    DO_PIECE = _doPiece,
                                    //DO_PRIX = decimal.Parse(txt_prix.Text),
                                    //DO_MONTANT = decimal.Parse(txt_montant.Text),
                                    //DO_POIDS = decimal.Parse(txt_poids.Text)
                                };

                                _context.F_FRETS.Add(f_fret);
                            }

                            _context.SaveChanges();


                            try
                            {
                                gridViewFrais.CloseEditor();
                                gridViewFrais.UpdateCurrentRow();

                                _context.SaveChanges();

                                // Exécuter la procédure
                                string sql = "EXEC SP_CalculCoutRevientParValeur @DO_Piece, @prix_total,@poids_total";

                                _context.Database.ExecuteSqlCommand(
                                   sql,
                                   new SqlParameter("@DO_Piece", _doPiece),
                                   new SqlParameter("@prix_total", Convert.ToDecimal(_parentForm.txt_prix.Text)),
                                   new SqlParameter("@poids_total", Convert.ToDecimal(_parentForm.txt_poids.Text))
                               );

                                // Rafraîchir la grille du parent
                                _parentForm?.InitializeGrid(_parentForm.GridLigneEdit, _doPiece);


                                MessageBox.Show("Frais enregistrés et coût de revient mis à jour avec succès.");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Erreur : " + ex.Message);
                            }

                        }
                        else
                        {
                            MessageBox.Show(
                                "Vous n'avez pas l'autorisation de modifier une facture !",
                                "Modification bloquée",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                    else if (_doPiece.ToString().StartsWith("ABC"))
                    {
                        bool autorise = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");

                        if (autorise)
                        {
                            //Enregistrer Fret

                            // Recherche si le DO_PIECE existe déjà
                            var existingFret = _context.F_FRETS
                                .FirstOrDefault(f => f.DO_PIECE == _doPiece);

                            if (existingFret != null)
                            {
                                //existingFret.DO_PRIX = decimal.Parse(txt_prix.Text);
                                //existingFret.DO_MONTANT = decimal.Parse(txt_montant.Text);
                                //existingFret.DO_POIDS = decimal.Parse(txt_poids.Text);
                            }
                            else
                            {
                                var f_fret = new F_FRET
                                {
                                    DO_PIECE = _doPiece,
                                    //DO_PRIX = decimal.Parse(txt_prix.Text),
                                    //DO_MONTANT = decimal.Parse(txt_montant.Text),
                                    //DO_POIDS = decimal.Parse(txt_poids.Text)
                                };

                                _context.F_FRETS.Add(f_fret);
                            }

                            _context.SaveChanges();


                            try
                            {
                                gridViewFrais.CloseEditor();
                                gridViewFrais.UpdateCurrentRow();

                                _context.SaveChanges();

                                // Exécuter la procédure
                                string sql = "EXEC SP_CalculCoutRevientParValeur @DO_Piece, @prix_total,@poids_total";

                                _context.Database.ExecuteSqlCommand(
                                    sql,
                                    new SqlParameter("@DO_Piece", _doPiece),
                                    new SqlParameter("@prix_total", Convert.ToDecimal(_parentForm.txt_prix.Text)),
                                    new SqlParameter("@poids_total", Convert.ToDecimal(_parentForm.txt_poids.Text))
                                );

                                // Rafraîchir la grille du parent
                                _parentForm?.InitializeGrid(_parentForm.GridLigneEdit, _doPiece);


                                MessageBox.Show("Frais enregistrés et coût de revient mis à jour avec succès.");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Erreur : " + ex.Message);
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
                }
            }
        }




        private void btnSupprimer_Click(object sender, EventArgs e)
        {
            if (tester_cloturer(_doPiece))
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
                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");
                bool autorise1 = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");
                bool autorise2 = frmMenuAchat.verifier_droit("Facture", "UPDATE");
                bool autorise3 = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                if (autorise || autorise1 || autorise2 || autorise3)
                {
                    var selectedRow = gridViewFrais.GetFocusedRow() as F_DOCFRAISIMPORT;
                    if (selectedRow != null)
                    {
                        _context.F_DOCFRAISIMPORT.Remove(selectedRow);
                        _bindingFrais.Remove(selectedRow);
                        _context.SaveChanges();
                    }
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de modifier ce document!",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private void btnAjouter_Click(object sender, EventArgs e)
        {
            if (tester_cloturer(_doPiece))
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
                bool autorise = frmMenuAchat.verifier_droit("Projet d'achat", "UPDATE");
                bool autorise1 = frmMenuAchat.verifier_droit("Bon de commande", "UPDATE");
                bool autorise2 = frmMenuAchat.verifier_droit("Facture", "UPDATE");
                bool autorise3 = frmMenuAchat.verifier_droit("Bon de réception", "UPDATE");

                if (autorise || autorise1 || autorise2 || autorise3)
                {
                    //txt_poids.Enabled = true;
                    //txt_prix.Enabled = true;

                    var newFrais = new F_DOCFRAISIMPORT
                    {
                        DO_Piece = _doPiece,
                        FI_Montant = 0,
                        FI_Montant_AR = 0,
                        FI_Devise = "MGA",           // ou autre par défaut
                        FI_TypeFraisId = 2,          // Assure-toi que l'ID 1 existe
                        FI_RepartitionId = 0,         // Idem, valeur par défaut valable
                        cbModification = DateTime.Now,
                        Username = FrmMdiParent._id_user,
                    };

                    _bindingFrais.Add(newFrais);
                    _context.F_DOCFRAISIMPORT.Add(newFrais);

                    // Placer le curseur sur la nouvelle ligne pour édition immédiate
                    var newRowHandle = gridViewFrais.RowCount - 1;
                    gridViewFrais.FocusedRowHandle = newRowHandle;
                    gridViewFrais.FocusedColumn = gridViewFrais.VisibleColumns[0];
                    gridViewFrais.ShowEditor();
                    gridViewFrais.Columns["cbModification"].OptionsColumn.ReadOnly = true;
                    gridViewFrais.Columns["Username"].OptionsColumn.ReadOnly = true;
                    gridViewFrais.Columns["Username"].VisibleIndex = -1;
                    gridViewFrais.RefreshData();
                    gridViewFrais.BestFitColumns();
                }
                else
                {
                    MessageBox.Show(
                        "Vous n'avez pas l'autorisation de modifier ce document!",
                        "Modification bloquée",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private void LoadData()
        {
            var devises = _context.F_DEVISES.ToList();

            if (devises.Any())
            {
                fDEVISEBindingSource.DataSource = devises;

                GridView gd_devise = new GridView(gridControlDevises);
                gridControlDevises.MainView = gd_devise;
                gridControlDevises.ViewCollection.Add(gd_devise);
                gridControlDevises.DataSource = fDEVISEBindingSource;

                gd_devise.PopulateColumns();
                gd_devise.Columns[0].Visible = false; // Masquer l'ID si nécessaire
                gd_devise.OptionsView.NewItemRowPosition = NewItemRowPosition.Top;
                gd_devise.OptionsBehavior.Editable = true;
            }
        }


        private void gridViewFrais_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
        {
            gridViewFrais.SetRowCellValue(e.RowHandle, "DO_Piece", _doPiece);
        }

        private void gridViewFrais_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName == "FI_TypeFraisId")
            {
                var currentValue = gridViewFrais.GetRowCellValue(e.RowHandle, e.Column);

                for (int i = 0; i < gridViewFrais.RowCount; i++)
                {
                    if (i != e.RowHandle) // ignorer la ligne actuelle
                    {
                        var otherValue = gridViewFrais.GetRowCellValue(i, e.Column);

                        if (object.Equals(currentValue, otherValue))
                        {
                            MessageBox.Show("Ce type de frais a déjà été sélectionné.", "Duplication interdite", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            // Annule la valeur saisie
                            gridViewFrais.SetRowCellValue(e.RowHandle, e.Column, null);
                            break;
                        }
                    }
                }

                var typeFrais = _context.P_TYPEFRAIS.FirstOrDefault(x => x.FR_Num == (int)currentValue);
                
                if (typeFrais != null && typeFrais.FR_Intitule == "TVA")
                {
                    gridViewFrais.OptionsBehavior.Editable = false;

                    try
                    {
                        //TotalTVA est une paramètre à l'appel de la formulaire depuis frmEditDocument
                        // on le trouve dans frmEditDocument barButtonItem5_ItemClick
                        gridViewFrais.SetRowCellValue(e.RowHandle, "FI_Montant", TotalTVA);
                    }
                    finally
                    {
                        gridViewFrais.OptionsBehavior.Editable = true;
                    }
                }
            }
        }


        //Calcul FRET
        private void txt_poids_EditValueChanged(object sender, EventArgs e)
        {
            //calcul_fret(txt_prix, txt_poids, txt_montant);
        }

        private void txt_prix_EditValueChanged(object sender, EventArgs e)
        {
            //calcul_fret(txt_prix,txt_poids,txt_montant);
        }

        private void calcul_fret(TextEdit t1,TextEdit t2,TextEdit t3)
        {
            /*gridViewFrais.SetRowCellValue(0, "FI_RepartitionId", 1);
            if (t1.Text != "" && t2.Text != "")
            {
                t3.Text = (Math.Round(double.Parse(t1.Text.ToString().Replace('.',','))  * double.Parse(t2.Text.ToString().Replace('.', ',')),2)).ToString() ;
            }
            else
            {
                t3.Text = "0";
            }

            for (int i = 0; i < gridViewFrais.RowCount; i++)
            {
                int rowHandle = gridViewFrais.GetVisibleRowHandle(i);

                if (gridViewFrais.IsDataRow(rowHandle))
                {
                    object cellValue = gridViewFrais.GetRowCellValue(rowHandle, "FI_TypeFraisId");

                    // Vérifie que la cellule contient une valeur convertible en entier
                    if (cellValue != null && int.TryParse(cellValue.ToString(), out int result))
                    {
                        if (result == 2)
                        {
                            // Vérifie que txt_montant contient une valeur numérique
                            if (decimal.TryParse(txt_montant.Text, out decimal montant) && montant != 0)
                            {
                                // Met à jour la cellule "FI_Montant" de la ligne courante
                                gridViewFrais.SetRowCellValue(rowHandle, "FI_Montant", montant);
                                break;
                            }
                        }
                    }
                }
            }

            // Rafraîchit le pied de page après la boucle
            gridViewFrais.RefreshData();
            gridViewFrais.InvalidateFooter();*/

        }

        private void txt_prix_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;

            // Autoriser chiffres, point, virgule, et touches de contrôle
            if (!char.IsControl(ch) && !char.IsDigit(ch) && ch != '.' && ch != ',')
            {
                e.Handled = true; // Bloque la saisie
            }

            // Empêcher plusieurs points ou virgules
            TextEdit te = sender as TextEdit;
            if ((ch == '.' || ch == ',') && (te.Text.Contains('.') || te.Text.Contains(',')))
            {
                e.Handled = true;
            }
        }

        private void txt_poids_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;

            // Autoriser chiffres, point, virgule, et touches de contrôle
            if (!char.IsControl(ch) && !char.IsDigit(ch) && ch != '.' && ch != ',')
            {
                e.Handled = true; // Bloque la saisie
            }

            // Empêcher plusieurs points ou virgules
            TextEdit te = sender as TextEdit;
            if ((ch == '.' || ch == ',') && (te.Text.Contains('.') || te.Text.Contains(',')))
            {
                e.Handled = true;
            }
        }

        private void frmEditInfos_Activated(object sender, EventArgs e)
        {
            //txt_montant.Enabled=false;
            gridViewFrais.CellValueChanged += GridViewFrais_CellValueChanged;
            gridViewFrais.Columns["FLAG1"].Visible = false;
            gridViewFrais.Columns["FLAG2"].Visible = false;
            gridViewFrais.Columns["FLAG3"].Visible = false;
            gridViewFrais.OptionsView.ShowFooter = true;
        }

        private void gridViewFrais_CustomDrawFooterCell(object sender, DevExpress.XtraGrid.Views.Grid.FooterCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName == "FI_Montant_AR")
            {
                // Récupérer le total du grid
                decimal total = 0;
                decimal tot = 0;
                if (gridViewFrais.Columns["FI_Montant_AR"].SummaryItem.SummaryValue != null)
                {
                    decimal.TryParse(gridViewFrais.Columns["FI_Montant_AR"].SummaryItem.SummaryValue.ToString(), out total);
                }

                // Récupérer la valeur du TextBox
                decimal txtMontant = 0;
                //decimal.TryParse(txt_montant.Text, out txtMontant);
                tot = 0;
                tot = total + txtMontant;

                // Affichage personnalisé
                e.Info.DisplayText = $"Total = {tot:n2}";
            }
        }

        private void GridViewFrais_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName == "FI_Devise" || e.Column.FieldName == "FI_Montant")
            {
                int rowHandle = e.RowHandle;

                object valeurDevise = gridViewFrais.GetRowCellValue(rowHandle, "FI_Devise");
                double montant = 0;
                object valeur_montant = gridViewFrais.GetRowCellValue(rowHandle, "FI_Montant");


                if (valeur_montant != null && double.TryParse(valeur_montant.ToString(), out double temp))
                {
                    montant = temp;
                }

                if (!valeurDevise.ToString().Contains("MGA"))
                {
                    double valeur_devise = recuperer_valeur_devise(valeurDevise.ToString());

                    decimal valeurDec = (decimal)(valeur_devise * montant);

                    gridViewFrais.SetRowCellValue(rowHandle, "FI_Montant_AR", valeurDec);
                }
                else
                {
                    decimal valeurDec = (decimal)montant;
                    gridViewFrais.SetRowCellValue(rowHandle, "FI_Montant_AR", valeurDec);
                }
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {

            GridView view = (GridView)gridControlDevises.MainView;
            
            for (int i = 0; i < view.RowCount; i++)
            {
                string devise = view.GetRowCellValue(i, "devise")?.ToString();
                string valeurStr = view.GetRowCellValue(i, "valeur")?.ToString();

                if (string.IsNullOrWhiteSpace(devise))
                    continue; // Ignore les lignes vides

                // Recherche de la devise existante
                var existingDevise = _context.F_DEVISES
                    .FirstOrDefault(d => d.devise == devise);

                if (existingDevise != null)
                {
                    // Mise à jour si elle existe
                    existingDevise.valeur = decimal.Parse(valeurStr.ToString());
                }
                else
                {
                    // Insertion sinon
                    var newDevise = new F_DEVISE
                    {
                        devise = devise,
                        valeur = decimal.Parse(valeurStr.ToString())
                    };
                    _context.F_DEVISES.Add(newDevise);
                }
            }

            _context.SaveChanges();
            LoadData();
        }

        private double recuperer_valeur_devise(string devise)
        {
            double val = 0;

            string connectionString = $"Data Source=26.53.123.231;Initial Catalog=ARBIOCHEM_ACHAT;User ID=Dev;Password=1234;TrustServerCertificate=True";

            string query = "SELECT valeur FROM F_DEVISE WHERE devise=@devise";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@devise", devise);
                object result = cmd.ExecuteScalar();

                if (result != null && result != DBNull.Value)
                {
                    double.TryParse(result.ToString(), out val);
                }
                conn.Close();
            }

            return val;
        }

        private bool tester_existence_devise(string devise)
        {
            bool test = false;

            string connectionString = $"Data Source=26.53.123.231;Initial Catalog=ARBIOCHEM;User ID=Dev;Password=1234;TrustServerCertificate=True";

            string query = "SELECT * FROM F_DEVISE WHERE devise=@devise";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@devise", devise);
                SqlDataReader reader = cmd.ExecuteReader();

                DataTable dt = new DataTable();
                dt.Load(reader);

                if (dt.Rows.Count > 0)
                {
                    test = true;
                }
                conn.Close();
            }

            return test;
        }
    }
}