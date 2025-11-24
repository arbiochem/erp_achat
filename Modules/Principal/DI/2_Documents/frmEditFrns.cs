using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using arbioApp.Modules.Principal.DI.Services;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Utils.Connection;
using DevExpress.ChartRangeControlClient.Core;
using DevExpress.DataAccess.DataFederation;
using DevExpress.XtraCharts.Designer.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Columns;
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

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frmEditFrns : DevExpress.XtraEditors.XtraForm
    {
        private string ctNum;
        
        public frmEditFrns(string frns, System.Windows.Forms.BindingSource sourceFrns)
        {
            InitializeComponent();
            this.Text= txtCtIntitule.Text;
            ctNum = frns;
            bindingNavigator1.BindingSource = sourceFrns;
            try
            {
                if (sourceFrns.DataSource is DataTable dataTableFrns)
                {

                    // Vérifier si la colonne CT_Num existe
                    if (!dataTableFrns.Columns.Contains("CT_Num"))
                    {
                        MessageBox.Show("La structure des données est incorrecte. Colonne CT_Num manquante.");
                        this.Close();
                        return;
                    }

                    // Trouver la position
                    int rowcount = dataTableFrns.Rows.Count;
                    int position = dataTableFrns.Rows.IndexOf(
                        dataTableFrns.AsEnumerable()
                        .FirstOrDefault(row => row.Field<string>("CT_Num") == frns)
                    );

                    if (position >= 0)
                    {
                        sourceFrns.Position = position;
                        bindingNavigator1.BindingSource = sourceFrns;
                        ChargerDepot();
                        ChargerDevise();
                        ChargerExpedition();
                        ChargerPays();
                        ChargerCompteCollectif();
                        LierChamps(sourceFrns);
                    }
                    else
                    {
                        MessageBox.Show($"Fournisseur {frns} non trouvé dans la liste.");
                        this.Close();
                    }
                }
                else
                {
                    MessageBox.Show("La source de données n'est pas un DataTable valide.");
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de l'ouverture du fournisseur : " + ex.Message);
                this.Close();
            }
        }
        List<F_DEPOT> _listeDepot;
        private void ChargerDepot()
        {
            _listeDepot = Entetes.GetAllDepots();

            lkCtDepot.Properties.DataSource = _listeDepot;
            lkCtDepot.Properties.ValueMember = "DE_No"; // Clé réelle stockée
            lkCtDepot.Properties.DisplayMember = "DE_Intitule"; // Texte affiché
            lkCtDepot.Properties.PopulateColumns();
            lkCtDepot.Properties.Columns.Clear();

            lkCtDepot.Properties.Columns.Add(new LookUpColumnInfo("DE_No", "DE_No", 50));
            lkCtDepot.Properties.Columns.Add(new LookUpColumnInfo("DE_Intitule", "DEPOT"));

        }

        List<P_DEVISE> _listeDevise;
        private void ChargerDevise()
        {
            _listeDevise = Entetes.GetAllDevise();

            lkCtDevise.Properties.DataSource = _listeDevise;

            lkCtDevise.Properties.ValueMember = "cbMarq";       // ✅ Clé primaire (utilisée en base)
            lkCtDevise.Properties.DisplayMember = "D_Intitule"; // ✅ Texte lisible (affiché)

            lkCtDevise.Properties.PopulateColumns();
            lkCtDevise.Properties.Columns.Clear();
            lkCtDevise.Properties.Columns.Add(new LookUpColumnInfo("cbMarq", "cbMarq",50));
            lkCtDevise.Properties.Columns.Add(new LookUpColumnInfo("D_Intitule", "DEVISE"));

            // Optionnel : empêcher la saisie manuelle
            lkCtDevise.Properties.TextEditStyle = TextEditStyles.Standard; // ou .DisableTextEditor
        }

        List<P_EXPEDITION> _listeExpedition;
        private void ChargerExpedition()
        {
            _listeExpedition = Entetes.GetAllExpedition();
            lkCtExpedition.Properties.DataSource = _listeExpedition;
            lkCtExpedition.Properties.DataSource = _listeExpedition;
            lkCtExpedition.Properties.ValueMember = "cbMarq";    // Clé réelle stockée
            lkCtExpedition.Properties.DisplayMember = "E_Intitule"; // Texte affiché
            lkCtExpedition.Properties.PopulateColumns();
            lkCtExpedition.Properties.Columns.Clear();

            lkCtExpedition.Properties.Columns.Add(new LookUpColumnInfo("cbMarq", "cbMarq",50));
            lkCtExpedition.Properties.Columns.Add(new LookUpColumnInfo("E_Intitule", "EXPEDITION"));

        }
        private  string dbPrincipale = ucDocuments.dbNamePrincipale;
        private  string serveripPrincipale = ucDocuments.serverIpPrincipale;
        private void ChargerPays()
        {
            string connectionString = $"Server={serveripPrincipale};" +
                                      $"Database={dbPrincipale};User ID=Dev;Password=1234;" +
                                      $"TrustServerCertificate=True;Connection Timeout=120;";

            string query = "SELECT DISTINCT CT_Pays FROM F_COMPTET WHERE ISNULL(CT_Pays, '') <> '' ORDER BY CT_Pays";

            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                try
                {
                    conn.Open();
                    adapter.Fill(dt);

                    lkCtPays.Properties.Items.Clear(); // Nettoie les anciens items

                    foreach (DataRow row in dt.Rows)
                    {
                        string pays = row["CT_Pays"].ToString();
                        lkCtPays.Properties.Items.Add(pays);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur lors du chargement des pays : " + ex.Message);
                }
            }
        }
        List<F_COMPTEG> _listeCompteG;
        private void ChargerCompteCollectif()
        {
            _listeCompteG = Entetes.GetAllCompteG();

            lkCptCol.Properties.DataSource = _listeCompteG;
            lkCptCol.Properties.ValueMember = "CG_Num"; // Clé réelle stockée
            lkCptCol.Properties.DisplayMember = "CG_Intitule"; // Texte affiché
            lkCptCol.Properties.PopulateColumns();
            lkCptCol.Properties.Columns.Clear();

            lkCptCol.Properties.Columns.Add(new LookUpColumnInfo("CG_Num", "CG_Num", 50));
            lkCptCol.Properties.Columns.Add(new LookUpColumnInfo("CG_Intitule", "CompteG"));

        }

        private void LierChamps(System.Windows.Forms.BindingSource bindingSource)
        {
            txtCtNum.DataBindings.Add("Text", bindingSource, "CT_Num");
            txtCtIntitule.DataBindings.Add("Text", bindingSource, "CT_Intitule");
            txtCtAbr.DataBindings.Add("Text", bindingSource, "CT_Classement");
            txtCtCommentaire.DataBindings.Add("Text", bindingSource, "CT_Commentaire");
            txtCtAdresse.DataBindings.Add("Text", bindingSource, "CT_Adresse");
            txtCtComplement.DataBindings.Add("Text", bindingSource, "CT_Complement");
            txtCtVille.DataBindings.Add("Text", bindingSource, "CT_Ville");
            txtCtCodePostal.DataBindings.Add("Text", bindingSource, "CT_CodePostal");
            txtCtRegion.DataBindings.Add("Text", bindingSource, "CT_CodeRegion");
            lkCtPays.DataBindings.Add("EditValue", bindingSource, "CT_Pays");
            txtCtTelephone.DataBindings.Add("Text", bindingSource, "CT_Telephone");
            txtCtTelecopie.DataBindings.Add("Text", bindingSource, "CT_Telecopie");
            txtCtLinkedin.DataBindings.Add("Text", bindingSource, "CT_LinkedIn");
            txtCtEmail.DataBindings.Add("Text", bindingSource, "CT_Email");
            txtCtFacebook.DataBindings.Add("Text", bindingSource, "CT_Facebook");
            txtCtInternet.DataBindings.Add("Text", bindingSource, "CT_Site");
            txtCtSiret.DataBindings.Add("Text", bindingSource, "CT_Siret");            
            lkCtDepot.DataBindings.Add("EditValue", bindingSource, "DE_No");
            lkCtDevise.DataBindings.Add("EditValue", bindingSource, "N_Devise");
            lkCtExpedition.DataBindings.Add("EditValue", bindingSource, "N_Expedition");
            lkCtConditionLivr.DataBindings.Add("EditValue", bindingSource, "N_Condition");
            lkCptCol.DataBindings.Add("EditValue", bindingSource, "CG_NumPrinc");
            chkFrnSommeil.DataBindings.Add("Checked", bindingSource, "CT_Sommeil", true, DataSourceUpdateMode.OnPropertyChanged);//

            //// Convertit le smallint (0/1) en bool (false/true)
            //chkFrnSommeil.DataBindings["Checked"].Format += (s, e) =>
            //{
            //    if (e.DesiredType == typeof(bool))
            //    {
            //        e.Value = ConvertFromDBVal<string>(e.Value != null && Convert.ToInt16(e.Value) == 1);
            //    }
            //};

            //// Convertit le bool (false/true) en smallint (0/1)
            //chkFrnSommeil.DataBindings["Checked"].Parse += (s, e) =>
            //{
            //    if (e.DesiredType == typeof(object))
            //    {
            //        e.Value = ((bool)e.Value) ? (short)1 : (short)0;
            //    }
            //};

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
        private void ResetFormFrns()
        {
            txtCtNum.Text = "";
            txtCtNum.Focus();
            txtCtIntitule.Text = "";
            txtCtAbr.Text = "";
            lkCptCol.EditValue = "";
            txtCtCommentaire.Text = "";
            txtCtAdresse.Text = "";
            txtCtComplement.Text = "";
            txtCtVille.Text = "";
            txtCtCodePostal.Text = "";
            txtCtRegion.Text = "";
            lkCtPays.EditValue = "";
            txtCtTelephone.Text = "";
            txtCtTelecopie.Text = "";
            txtCtLinkedin.Text = "";
            txtCtEmail.Text = "";
            txtCtFacebook.Text = "";
            txtCtInternet.Text = "";
            txtCtSiret.Text = "";
            txtCtCodeNaf.Text = "";
            txtCtIdTVA.Text = "";
            lkCtDepot.Text = "";
            lkCtCodeAffaire.Text = "";
            lkCtDevise.Text = "";
            lkCtExpedition.Text = "";
            lkCtConditionLivr.Text = "";
        }
        private void btrnNouvCt_Click(object sender, EventArgs e)
        {
            txtCtNum.Properties.ReadOnly = false;
            ResetFormFrns();
        }
        public event EventHandler DonneesMisesAJour;
        protected virtual void OnDonneesMisesAJour()
        {
            DonneesMisesAJour?.Invoke(this, EventArgs.Empty);
        }
        private void btnOkCt_Click(object sender, EventArgs e)
        {
            try
            {
                short sommeil = chkFrnSommeil.Checked ? (short)1 : (short)0;
                var compte = new F_COMPTET
                {
                    CT_Num = txtCtNum.Text.Trim(),
                    CT_Intitule = txtCtIntitule.Text.Trim(),
                    CT_Type = 1,
                    CT_Qualite = "FOURNISSEUR",
                    CT_Classement = txtCtAbr.Text.Trim(),
                    CT_Adresse = txtCtAdresse.Text.Trim(),
                    CT_Complement = txtCtComplement.Text.Trim(),
                    CT_CodePostal = txtCtCodePostal.Text.Trim(),
                    CT_Ville = txtCtVille.Text.Trim(),
                    CT_Pays = lkCtPays.Text.Trim(),
                    CT_CodeRegion = txtCtRegion.Text.Trim(),
                    CT_Identifiant = txtCtIdTVA.Text.Trim(),
                    CT_Siret = txtCtSiret.Text.Trim(),
                    CT_Commentaire = txtCtCommentaire.Text.Trim(),
                    CT_NumPayeur = txtCtNum.Text.Trim(),
                    CT_Sommeil = sommeil,
                    CT_Telephone = txtCtTelephone.Text.Trim(),
                    CT_Telecopie = txtCtTelecopie.Text.Trim(),
                    CT_Facebook= txtCtFacebook.Text.Trim(),
                    CT_LinkedIn = txtCtLinkedin.Text.Trim(),
                    CT_EMail = txtCtEmail.Text.Trim(),
                    CT_Site = txtCtInternet.Text.Trim(),
                    CG_NumPrinc = lkCptCol.EditValue != null ? lkCptCol.EditValue.ToString() : null,
                    // ✅ Ici on force 0 si null
                    DE_No = lkCtDepot.EditValue != null ? Convert.ToInt32(lkCtDepot.EditValue) : 0,
                    N_Devise = lkCtDevise.EditValue != null ? Convert.ToInt16(lkCtDevise.EditValue) : (short)0,
                    N_Expedition = lkCtExpedition.EditValue != null ? (short?)Convert.ToInt32(lkCtExpedition.EditValue) : (short?)0,
                    CT_DateMAJ = DateTime.Now
                };


                var service = new F_COMPTETService(new F_COMPTETRepository(new AppDbContext()));
                service.SaveOrUpdate(compte);

                MessageBox.Show("Enregistrement effectué avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetFormFrns();
                OnDonneesMisesAJour();
                txtCtNum.Properties.ReadOnly = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'enregistrement : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void frmEditFrns_Load_1(object sender, EventArgs e)
        {
            bindingNavigator1.BindingSource.PositionChanged += BindingSource_PositionChanged;
            txtCtNum.Properties.ReadOnly = true; // Rendre le champ CT_Num en lecture seule
        }
        private void BindingSource_PositionChanged(object sender, EventArgs e)
        {
            this.Text = txtCtIntitule.Text;
        }

       
    }
}