using arbioApp.DTO;
using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using arbioApp.Modules.Principal.DI.Services;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Services;
using arbioApp.Utils.Connection;
using DevExpress.ChartRangeControlClient.Core;
using DevExpress.DataAccess.DataFederation;
using DevExpress.DataAccess.Native.Data;
using DevExpress.XtraCharts.Designer.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraSpreadsheet.DocumentFormats.Xlsb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmEditArticles : DevExpress.XtraEditors.XtraForm
    {
        private readonly AppDbContext _context;
        private string ArRef;
        private  F_ARTFOURNISSService _artFournissService;
        private BindingList<F_ARTFOURNISS> _dataSource;
        public frmEditArticles(string Ar_ref, System.Windows.Forms.BindingSource sourceArt)
        {
            InitializeComponent();
            this.Text= Ar_ref;
            ArRef = Ar_ref;
            bindingNavigator1.BindingSource = sourceArt;
            try
            {
                var article = sourceArt.List
                                 .OfType<F_ARTICLE>()
                                 .FirstOrDefault(a => a.AR_Ref == ArRef);

                int rowcount = sourceArt.Count;
                int position = sourceArt.Position;

                    if (position >= 0)
                    {
                        sourceArt.Position = position;
                        ChargerUnite();

                        LierChamps(sourceArt);
                    }
                    else
                    {
                        MessageBox.Show($"Fournisseur {ArRef} non trouvé dans la liste.");
                        this.Close();
                    }

                InitializeGrid();
                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de l'ouverture du fournisseur : " + ex.Message);
                this.Close();
            }
        }
        List<P_UNITE> _listeUnite;
        private void ChargerUnite()
        {
            _listeUnite = Entetes.GetAllUnites();

            lkArUnite.Properties.DataSource = _listeUnite;
            lkArUnite.Properties.ValueMember = "cbIndice"; // Clé réelle stockée
            lkArUnite.Properties.DisplayMember = "U_Intitule"; // Texte affiché
            lkArUnite.Properties.PopulateColumns();
            lkArUnite.Properties.Columns.Clear();

            lkArUnite.Properties.Columns.Add(new LookUpColumnInfo("cbIndice", "cbIndice", 50));
            lkArUnite.Properties.Columns.Add(new LookUpColumnInfo("U_Intitule", "Unité"));

        }
        private void LierChamps(System.Windows.Forms.BindingSource bindingSource)
        {
            txtArRef.DataBindings.Add("Text", bindingSource, "AR_Ref");
            txtArDesign.DataBindings.Add("Text", bindingSource, "AR_Design");
            txtArFa.DataBindings.Add("Text", bindingSource, "FA_CodeFamille");
            lkArUnite.DataBindings.Add("EditValue", bindingSource, "AR_UniteVen", true, DataSourceUpdateMode.OnPropertyChanged);

            txtArPA.DataBindings.Add("EditValue", bindingSource, "AR_PrixAch", true, DataSourceUpdateMode.OnPropertyChanged, 0);
            txtArDPA.DataBindings.Add("EditValue", bindingSource, "AR_PUNet", true, DataSourceUpdateMode.OnPropertyChanged, 0);
            //spQteMini.DataBindings.Add("EditValue", bindingSource, "artStock.AS_QteMini", true, DataSourceUpdateMode.OnPropertyChanged, 0);
            //spQteMaxi.DataBindings.Add("EditValue", bindingSource, "artStock.AS_QteMaxi", true, DataSourceUpdateMode.OnPropertyChanged, 0);
        }



        public static T ConvertFromDBVal<T>(object obj)
        {
            if (obj == null || obj == DBNull.Value)
            {
                return default(T); // returns the default value for the type
            }
            else
            {
                return (T)obj;
            }
        }
       
        private void btrnNouvCt_Click(object sender, EventArgs e)
        {
            txtArRef.Properties.ReadOnly = false;
            //ResetFormArt();
        }
       
        private void btnOkCt_Click(object sender, EventArgs e)
        {
            try
            {
                var article = (F_ARTICLE)((System.Windows.Forms.BindingSource)bindingNavigator1.BindingSource).Current;
                if (article != null)
                {
                    // Les bindings ont déjà mis à jour article.AR_PrixAch, AR_PUNet, AR_UniteVen etc.
                    // Si tu veux forcer une date de modif :
                    article.AR_DateModif = DateTime.Now;

                    // Si F_ARTSTOCK est chargé en navigation :
                    if (article.artStock != null)
                    {
                        //article.artStock.AS_QteMini = ConvertFromDBVal<decimal>(spQteMini.EditValue);
                        //article.artStock.AS_QteMaxi = ConvertFromDBVal<decimal>(spQteMaxi.EditValue);
                        // AS_QteSto = lecture seule, on ne touche pas
                    }

                    using (var ctx = new AppDbContext())
                    {
                        var dbArticle = ctx.F_ARTICLE
                       .Include(a => a.artStock)
                       .FirstOrDefault(a => a.AR_Ref == article.AR_Ref);
                        if (dbArticle != null)
                        {
                            // Copie des nouvelles valeurs depuis ton "article" détaché
                            ctx.Entry(dbArticle).CurrentValues.SetValues(article);

                            if (article.artStock != null)
                            {
                                ctx.Entry(dbArticle.artStock).CurrentValues.SetValues(article.artStock);
                            }

                            ctx.SaveChanges();
                        }
                    }

                    MessageBox.Show("Enregistrement effectué avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


                var service = new F_ARTICLEService(new F_ARTICLERepository(new AppDbContext()));
                service.UpdateDateModifArticle(
                            ArRef,
                            ConvertFromDBVal<decimal>(txtArPA.EditValue),
                            ConvertFromDBVal<decimal>(txtArDPA.EditValue),
                            ConvertFromDBVal<short>(lkArUnite.EditValue)
                            //ConvertFromDBVal<decimal>(spQteMini.EditValue),
                            //ConvertFromDBVal<decimal>(spQteMaxi.EditValue)
                        );

                //service.UpdateDateModifArticle(ArRef, ConvertFromDBVal<decimal>(txtArPA.EditValue), ConvertFromDBVal<decimal>(txtArDPA.EditValue), ConvertFromDBVal<short>(lkArUnite.EditValue));

                //MessageBox.Show("Enregistrement effectué avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //txtArRef.Properties.ReadOnly = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'enregistrement : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        

        private void frmEditArticles_Load_1(object sender, EventArgs e)
        {
            bindingNavigator1.BindingSource.PositionChanged += BindingSource_PositionChanged;
            txtArRef.Properties.ReadOnly = true; // Rendre le champ CT_Num en lecture seule

            ChargerFournisseurs(ArRef);
            gvArFrns.Columns["AR_Ref"].Visible = true;
            gvArFrns.Columns["CT_Num"].Visible = true;
            gvArFrns.Columns["CT_Intitule"].Visible = true;
            gvArFrns.OptionsBehavior.Editable = true;
            gvArFrns.BestFitColumns();

        }
        private void BindingSource_PositionChanged(object sender, EventArgs e)
        {
            this.Text = txtArDesign.Text;
        }
        private System.Windows.Forms.BindingSource sourceArtFrns;
        private readonly F_ARTFOURNISSService _f_ARTFOURNISSService;
        private void ChargerFournisseurs(string arRef)
        {
            using (var ctx = new AppDbContext())
            {
                F_ARTFOURNISS fournisseur = _f_ARTFOURNISSService.GetByARRefAndPrincipal(ArRef);

                sourceArtFrns = new System.Windows.Forms.BindingSource { DataSource = fournisseur };
                gcArFrns.DataSource = sourceArtFrns;
            }
        }
        
        private void BtnAddFrns_Click(object sender, EventArgs e)
        {
            try
            {
                var newRecord = new F_ARTFOURNISS
                {
                    AR_Ref = string.Empty,
                    CT_Num = string.Empty,
                    CT_Intitule = string.Empty,
                    AF_Principal = 0 // Valeur par défaut
                };

                _dataSource.Add(newRecord);

                // Positionner sur la nouvelle ligne
                GridView view = gcArFrns.MainView as GridView;
                view.FocusedRowHandle = view.RowCount - 1;
                view.ShowEditor();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show($"Erreur lors de l'ajout: {ex.Message}", "Erreur");
            }
        }

        private void btnSaveFrns_Click(object sender, EventArgs e)
        {
            try
            {
                // Forcer la fin de l'édition en cours
                GridView view = gcArFrns.MainView as GridView;
                view.CloseEditor();
                view.UpdateCurrentRow();

                // Valider les données
                var invalidRecords = _dataSource.Where(x =>
                    string.IsNullOrWhiteSpace(x.AR_Ref) ||
                    string.IsNullOrWhiteSpace(x.CT_Num)).ToList();

                if (invalidRecords.Any())
                {
                    XtraMessageBox.Show("Veuillez remplir tous les champs obligatoires (AR_Ref et CT_Num)", "Validation");
                    return;
                }

                // Sauvegarder en base
                _artFournissService.SaveChanges(_dataSource.ToList());

                XtraMessageBox.Show("Données enregistrées avec succès!", "Succès");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show($"Erreur lors de l'enregistrement: {ex.Message}", "Erreur");
            }
        }
        private void InitializeGrid()
        {
            GridView view = gcArFrns.MainView as GridView;

            // Configuration du GridView
            view.OptionsBehavior.Editable = true;
            view.OptionsView.NewItemRowPosition = NewItemRowPosition.None; // Nous gérons l'ajout via bouton

            // Création des colonnes
            view.Columns.Clear();

            // Colonne AR_Ref
            GridColumn colAR_Ref = new GridColumn();
            colAR_Ref.FieldName = "AR_Ref";
            colAR_Ref.Caption = "Référence Article";
            colAR_Ref.Visible = true;
            colAR_Ref.OptionsColumn.AllowEdit = true;
            view.Columns.Add(colAR_Ref);

            // Colonne CT_Num
            GridColumn colCT_Num = new GridColumn();
            colCT_Num.FieldName = "CT_Num";
            colCT_Num.Caption = "Numéro Fournisseur";
            colCT_Num.Visible = true;
            colCT_Num.OptionsColumn.AllowEdit = true;
            view.Columns.Add(colCT_Num);

            // Colonne CT_Intitule
            GridColumn colCT_Intitule = new GridColumn();
            colCT_Intitule.FieldName = "CT_Intitule";
            colCT_Intitule.Caption = "Intitulé Fournisseur";
            colCT_Intitule.Visible = true;
            colCT_Intitule.OptionsColumn.AllowEdit = true;
            view.Columns.Add(colCT_Intitule);

            // Configuration supplémentaire
            view.BestFitColumns();
        }
       
        
    }
}