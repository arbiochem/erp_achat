using arbioApp.Models;
using DevExpress.Data.ODataLinq.Helpers;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frm_saisieQuantite : Form
    {
        // Propriétés publiques pour recevoir les valeurs depuis le formulaire appelant
        public string _var1 { get; set; }
        public string _var2 { get; set; }
        public string _num { get; set; }
        public string _fournisseur { get; set; }
        private string connectionStrings = "Server=26.53.123.231;Database=TRANSIT;User Id=Dev;Password=1234;";
        private readonly List<F_DEPOT> _listeDepot;
        public int deno;
        public int cono;
        public DateTime dt;
        private readonly AppDbContext _context;

        // Propriété pour récupérer la quantité saisie
        public decimal QuantiteSaisie { get; private set; }

        public frm_saisieQuantite()
        {
            InitializeComponent();
        }

        /*protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; // bloquer la fermeture par l'utilisateur
            }
            else
            {
                base.OnFormClosing(e);
            }
        }*/

        // Constructeur surchargé pour passer les valeurs directement
        public frm_saisieQuantite(string ref_, string design, string num,string fournisseur)
        {
            InitializeComponent();
            _var1 = ref_;
            _var2 = design;
            _num = num;
            _fournisseur=fournisseur;
        }

        private void frm_saisieQuantite_Load(object sender, EventArgs e)
        {
            afficher();
            lblref1.Text = _var1;
            lbl_design2.Text = _var2;
            lblnumdoc.Text = _num.ToString();
            lblFournisseur.Text = _fournisseur.ToString();
            ChargerDepot();
            ChargerListeAcheteur();
        }

        private void ChargerDepot()
        {
            List<F_DEPOT> _listeDepot;
            
            _listeDepot = Entetes.GetAllDepots();

            cmbdepot.Properties.DataSource = _listeDepot;
            cmbdepot.Properties.ValueMember = "DE_No"; // Clé réelle stockée
            cmbdepot.Properties.DisplayMember = "DE_Intitule"; // Texte affiché
            cmbdepot.Properties.PopulateColumns();
            cmbdepot.Properties.Columns.Clear();

            cmbdepot.Properties.Columns.Add(new LookUpColumnInfo("DE_No", "DE_No", 50));
            cmbdepot.Properties.Columns.Add(new LookUpColumnInfo("DE_Intitule", "DEPOT"));

        }

        public List<F_COMPTET> GetAllFournisseurs()
        {
            string query = "SELECT CT_Num, CT_Intitule FROM F_COMPTET WHERE CT_Type = 1";
            List<F_COMPTET> fournisseurs = new List<F_COMPTET>();

            using (SqlConnection conn = new SqlConnection(connectionStrings))
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

        private void ChargerListeAcheteur()
        {
            List<F_COLLABORATEUR> _listeAcheteurs;
            _listeAcheteurs = GetAllAcheteurs();

            cmbcollaborateur.Properties.DataSource = _listeAcheteurs;
            cmbcollaborateur.Properties.ValueMember = "CO_No"; // Clé réelle stockée
            cmbcollaborateur.Properties.DisplayMember = "CO_Nom"; // Texte affiché
            cmbcollaborateur.Properties.PopulateColumns();
            cmbcollaborateur.Properties.Columns.Clear();

            cmbcollaborateur.Properties.Columns.Add(new LookUpColumnInfo("CO_No", "CO_No", 50));
            cmbcollaborateur.Properties.Columns.Add(new LookUpColumnInfo("CO_Nom", "Acheteur"));

        }

        public List<F_COLLABORATEUR> GetAllAcheteurs()
        {
            List<F_COLLABORATEUR> Acheteurs = new List<F_COLLABORATEUR>();
            string query = "SELECT CO_No, CO_Nom + ' ' + CO_Prenom AS CO_Nom FROM F_COLLABORATEUR WHERE CO_Acheteur = 1";


            using (SqlConnection conn = new SqlConnection(connectionStrings))
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
        private void afficher()
        {
            try
            {
                string query = "SELECT DO_Cours, CO_No, DE_No, DO_DateLivr FROM F_DOCENTETE WHERE DO_Piece = @doPiece";
                using (var connection = new SqlConnection(connectionStrings))
                {
                    connection.Open();
                    using (var cmd = new SqlCommand(query, connection))
                    {
                        cmd.Parameters.Add("@doPiece", SqlDbType.Char, 13).Value = _num.PadRight(13);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // ✅ Récupérer les deux colonnes
                                decimal cours = reader["DO_Cours"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["DO_Cours"])
                                    : 0;

                                int coNo = reader["CO_No"] != DBNull.Value
                                    ? Convert.ToInt32(reader["CO_No"])
                                    : 0;

                                int deno = reader["DE_No"] != DBNull.Value
                                   ? Convert.ToInt32(reader["DE_No"])
                                   : 0;

                                txtcoursdevise.Text = cours.ToString();
                                cmbcollaborateur.EditValue = coNo == 0 ? null : (object)coNo;
                                cmbdepot.EditValue = deno == 0 ? null : (object)deno;
                                datelivrprev.EditValue = reader["DO_DateLivr"] != DBNull.Value
                                 ? Convert.ToDateTime(reader["DO_DateLivr"])
                                 : (object)null;
                            }
                            else
                            {
                                txtcoursdevise.Text = "0";
                                cmbcollaborateur.EditValue = null;
                                cmbdepot.EditValue = null;
                                datelivrprev.EditValue = null;
                            }
                        }
                    }
                }

                string query_fret = "SELECT * FROM F_FRET WHERE DO_PIECE = @doPiece";
                using (var connection = new SqlConnection(connectionStrings))
                {
                    connection.Open();
                    using (var cmd = new SqlCommand(query_fret, connection))
                    {
                        cmd.Parameters.Add("@doPiece", SqlDbType.Char, 13).Value = _num.PadRight(13);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // ✅ Récupérer les deux colonnes
                                decimal montant = reader["DO_Montant"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["DO_Montant"])
                                    : 0;

                                decimal poids = reader["DO_Poids"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["DO_Poids"])
                                    : 0;


                                txtFret.Text = montant.ToString();
                                txtPV.Text = poids.ToString();
                            }
                            else
                            {
                                txtFret.Text = "";
                                txtPV.Text = "";
                            }
                        }
                    }
                }
            }catch(Exception er) { }
        }

        private void txtQte_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Autoriser : chiffres et backspace uniquement
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }


        // Bouton Annuler
        private void btnAnnuler_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void enregistrer_fret()
        {
            try
            {
                using (AppDbContext context = new AppDbContext())
                {
                    // Pour un AJOUT (vérifier que l'enregistrement n'existe pas déjà)
                    var existingFRET = context.F_FRETS
                        .FirstOrDefault(s => s.DO_PIECE == _num.Trim());

                    if (existingFRET == null)
                    {
                        try
                        {
                            F_FRET f = new F_FRET();
                            f.DO_PIECE = _num.Trim();
                            f.DO_MONTANT = Convert.ToDecimal(txtFret.Text);
                            f.DO_POIDS = Convert.ToDecimal(txtPV.Text);

                            context.F_FRETS.Add(f);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    else
                    {
                        existingFRET.DO_MONTANT = Convert.ToDecimal(txtFret.Text);
                        existingFRET.DO_POIDS = Convert.ToDecimal(txtPV.Text);
                    }

                    context.SaveChanges();
                    //MessageBox.Show("Modification FRET terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Data.Entity.Infrastructure.DbUpdateException ex)
            {
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Erreur");
            }
        }

        private void btnValider_Click(object sender, EventArgs e)
        {
            enregistrer_fret();
            string texte = txtQte.Text.Replace(',', '.');

            if (decimal.TryParse(texte, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out decimal valeur))
            {
                QuantiteSaisie = valeur;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Veuillez saisir une quantité valide.", "Erreur",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtQte.Focus();
            }
        }

        private void txtKg_EditValueChanged(object sender, EventArgs e)
        {
            if (txtKg.Text != "")
            {
                decimal prixKg = decimal.Parse(txtKg.Text.ToString().Replace('.', ','));
                txtTonne.Text = Convert.ToString(prixKg * 1000);
            }
        }

        private void txtTonne_EditValueChanged(object sender, EventArgs e)
        {
            decimal prixTonne = decimal.Parse(txtTonne.Text.ToString());
            txtKg.Text = Convert.ToString(prixTonne / 1000);
        }

        private void txtKg_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string texteActuel = txtKg.Text;

            // Autoriser : chiffres, backspace
            bool estChiffre = char.IsDigit(c);
            bool estBackspace = c == '\b';

            // Autoriser virgule ou point UNE seule fois
            bool estVirgule = c == ',' && !texteActuel.Contains(",") && !texteActuel.Contains(".");
            bool estPoint = c == '.' && !texteActuel.Contains(",") && !texteActuel.Contains(".");

            // Bloquer tout le reste
            if (!estChiffre && !estBackspace && !estVirgule && !estPoint)
            {
                e.Handled = true; // bloquer la touche
            }
        }

        private void txtTonne_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string texteActuel = txtTonne.Text;

            // Autoriser : chiffres, backspace
            bool estChiffre = char.IsDigit(c);
            bool estBackspace = c == '\b';

            // Autoriser virgule ou point UNE seule fois
            bool estVirgule = c == ',' && !texteActuel.Contains(",") && !texteActuel.Contains(".");
            bool estPoint = c == '.' && !texteActuel.Contains(",") && !texteActuel.Contains(".");

            // Bloquer tout le reste
            if (!estChiffre && !estBackspace && !estVirgule && !estPoint)
            {
                e.Handled = true; // bloquer la touche
            }
        }

        private void cmbcollaborateur_EditValueChanged(object sender, EventArgs e)
        {
            if (cmbcollaborateur.EditValue != null && cmbcollaborateur.EditValue != DBNull.Value)
            {
                cono = Convert.ToInt32(cmbcollaborateur.EditValue);
            }
        }

        private void cmbdepot_EditValueChanged(object sender, EventArgs e)
        {
            if (cmbdepot.EditValue != null && cmbdepot.EditValue != DBNull.Value)
            {
                deno = Convert.ToInt32(cmbdepot.EditValue);
            }
        }

        private void datelivrprev_EditValueChanged(object sender, EventArgs e)
        {
            dt = Convert.ToDateTime(datelivrprev.EditValue.ToString());
        }

        private bool _isValidating = false;

        private void txtPoids_EditValueChanged(object sender, EventArgs e)
        {
            if (_isValidating) return; // ✅ Bloquer la récursion

            decimal poids_total = 0;

            if (string.IsNullOrEmpty(txtPoids.Text)) return;
            if (!decimal.TryParse(txtPoids.Text, out decimal poidsActuel) || poidsActuel <= 0) return;

            try
            {
                string query = "SELECT ISNULL(SUM(DL_PoidsNet), 0) FROM F_DOCLIGNE WHERE DO_Piece = @doPiece";
                using (var connection = new SqlConnection(connectionStrings))
                {
                    connection.Open();
                    using (var cmd = new SqlCommand(query, connection))
                    {
                        cmd.Parameters.Add("@doPiece", SqlDbType.Char, 13).Value = _num.PadRight(13);
                        object result = cmd.ExecuteScalar();
                        decimal sommeBD = result != null && result != DBNull.Value
                            ? Convert.ToDecimal(result) : 0;
                        poids_total = sommeBD + poidsActuel;
                    }
                }
            }
            catch { return; }

            if (!decimal.TryParse(txtPV.Text, out decimal poidsVolume)) return;

            if (poids_total > poidsVolume)
            {
                _isValidating = true; // ✅ Bloquer avant de vider
                try
                {
                    MessageBox.Show("Le poids total ne doit pas dépasser le poids total volume",
                        "Message d'erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtPoids.Text = "";
                    txtPoids.Refresh();
                }
                finally
                {
                    _isValidating = false; // ✅ Toujours réactiver
                }
            }
        }

        private void txtFret_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string texteActuel = txtKg.Text;

            // Autoriser : chiffres, backspace
            bool estChiffre = char.IsDigit(c);
            bool estBackspace = c == '\b';

            // Autoriser virgule ou point UNE seule fois
            bool estVirgule = c == ',' && !texteActuel.Contains(",") && !texteActuel.Contains(".");
            bool estPoint = c == '.' && !texteActuel.Contains(",") && !texteActuel.Contains(".");

            // Bloquer tout le reste
            if (!estChiffre && !estBackspace && !estVirgule && !estPoint)
            {
                e.Handled = true; // bloquer la touche
            }
        }

        private void txtPV_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string texteActuel = txtKg.Text;

            // Autoriser : chiffres, backspace
            bool estChiffre = char.IsDigit(c);
            bool estBackspace = c == '\b';

            // Autoriser virgule ou point UNE seule fois
            bool estVirgule = c == ',' && !texteActuel.Contains(",") && !texteActuel.Contains(".");
            bool estPoint = c == '.' && !texteActuel.Contains(",") && !texteActuel.Contains(".");

            // Bloquer tout le reste
            if (!estChiffre && !estBackspace && !estVirgule && !estPoint)
            {
                e.Handled = true; // bloquer la touche
            }
        }

        private void txtPoids_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string texteActuel = txtKg.Text;

            // Autoriser : chiffres, backspace
            bool estChiffre = char.IsDigit(c);
            bool estBackspace = c == '\b';

            // Autoriser virgule ou point UNE seule fois
            bool estVirgule = c == ',' && !texteActuel.Contains(",") && !texteActuel.Contains(".");
            bool estPoint = c == '.' && !texteActuel.Contains(",") && !texteActuel.Contains(".");

            // Bloquer tout le reste
            if (!estChiffre && !estBackspace && !estVirgule && !estPoint)
            {
                e.Handled = true; // bloquer la touche
            }
        }
    }
}