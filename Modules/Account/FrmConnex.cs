using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.CodeParser;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraLayout.Customization;
using DevExpress.XtraSplashScreen;
using Microsoft.Win32;
using System.Data.Sql;
using System.Data.SqlClient;
using System.IO;
using System.Security.Cryptography;
using arbioApp.Modules.Helpers;
using DevExpress.LookAndFeel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using arbioApp.Modules.Account;
using System.Reflection;


namespace arbioApp
{
    public partial class FrmConnex : DevExpress.XtraEditors.XtraForm
    {
        public string Username { get; private set; }
        public string IDName { get; private set; }
        public string UserRole { get; private set; }

        public bool success;
        public static string leserveur { get; set; }
        public string labase { get; set; }

        public string pwd
        {
            get; set;
        }
        public FrmConnex()
        {
            InitializeComponent();
            Addserver.LoadServerConnex(cboServers);
            //cboServers.SelectedIndex = 0;
            cboServers.EditValueChanged += cboServers_EditValueChanged;
        }
        private string GetSelectedThemeFromDatabase()
        {
            string selectedTheme = null;

            using (SqlConnection connection = new SqlConnection(
                       $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True"))
            {
                string query = "SELECT intitule_theme FROM T_Theme WHERE lOption = 1";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();

                    // Récupérer le nom du thème
                    object result = command.ExecuteScalar();
                    if (result != null)
                    {
                        selectedTheme = result.ToString();
                    }
                }
            }

            return selectedTheme;
        }
        private void ApplySelectedTheme()
        {
            // Récupérer le thème coché dans la base de données
            string selectedTheme = GetSelectedThemeFromDatabase();

            if (!string.IsNullOrEmpty(selectedTheme))
            {
                // Appliquer le thème récupéré
                UserLookAndFeel.Default.SetSkinStyle(selectedTheme);
            }
        }

        private void cboServers_EditValueChanged(object sender, EventArgs e)
        {
            BtnOK.Enabled = !string.IsNullOrEmpty(cboServers.EditValue?.ToString());
        }


        private void BtnOK_Click(object sender, EventArgs e)
        {
            leserveur = cboServers.EditValue?.ToString();
            pwd = textPwd.Text;
                testerAutorisation();
                this.DialogResult = DialogResult.OK;
        }

        public static string mailuser = "";
        private void testerAutorisation()
        {
            try
            {
                if (tokenEmail.EditValue.ToString() == "" || tokenEmail.EditValue == null)
                {
                    return;
                }
                string userName = tokenEmail.EditValue.ToString();
                string password = textPwd.Text;
                mailuser = tokenEmail.EditValue.ToString();
                if (string.IsNullOrWhiteSpace(userName) || string.IsNullOrWhiteSpace(password))
                {
                    MessageBox.Show("Veuillez entrer un nom d'utilisateur et un mot de passe.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (VerifyLogin(userName, password))
                {
                    using (SqlConnection connection = new SqlConnection(
                               $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True"))
                    {
                        try
                        {
                            connection.Open();

                            // Vérifier si l'utilisateur a changé son mot de passe
                            string query = "SELECT id_user, UserGroup, IDName, IsPasswordChanged FROM T_UserRole WHERE UserName = @UserName";
                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@UserName", userName);

                                using (SqlDataReader reader = command.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        Guid id_user = Guid.Parse(reader["id_user"].ToString());
                                        string userRole = reader["UserGroup"].ToString();
                                        string idname = reader["IDName"].ToString();
                                        bool isPasswordChanged = (bool)reader["IsPasswordChanged"];

                                        Username = userName;
                                        UserRole = userRole;
                                        IDName = idname;

                                        // Si le mot de passe n'a pas été changé, demander un changement de mot de passe
                                        if (!isPasswordChanged)
                                        {
                                            MessageBox.Show("Vous devez changer votre mot de passe pour la première connexion.");
                                            ShowChangePasswordForm(userName);
                                            return; // Empêche l'accès à l'application si le mot de passe n'est pas changé
                                        }

                                        DateHelper.ApposerLastDate(serverName, userName);

                                        FrmMdiParent mdiParent = new FrmMdiParent(id_user, idname, userName, userRole, serverName, "arbapp");
                                        ApplySelectedTheme();
                                        this.Hide();
                                        mdiParent.ShowDialog();
                                        mdiParent.BringToFront();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Utilisateur introuvable.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                }
                else
                {
                    MessageBox.Show("Nom d'utilisateur ou mot de passe incorrect, ou vérifiez le service réseau.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        // Fonction pour afficher le formulaire de changement de mot de passe
        private void ShowChangePasswordForm(string userName)
        {
            // Créer un formulaire pour changer le mot de passe
            FrmChangePassword changePasswordForm = new FrmChangePassword(userName);
            changePasswordForm.ShowDialog();
        }

        

        private bool VerifyLogin(string username, string inputPassword)
        {
            using (SqlConnection connection = new SqlConnection(
                       $"Data Source={serverName.Replace(";AMBOHIMANGAKELY", "")};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True"))
            {
                try
                {
                    connection.Open();

                    string query = "SELECT Password FROM T_UserRole WHERE Username = @UserName";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@UserName", username);

                        object result = command.ExecuteScalar();

                        if (result != null)
                        {
                            string storedHashedPassword = result.ToString();

                            string hashedInputPassword = PasswordHelper.HashPassword(inputPassword);

                            return hashedInputPassword == storedHashedPassword;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
        }


        private void FrmConnex_Load(object sender, EventArgs e)
        {           
            arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(tokenEmail);
        }
        string serverName;

        private void cboServers_EditValueChanged_1(object sender, EventArgs e)
        {
            serverName = cboServers.EditValue?.ToString();
            cboServers.SelectedIndexChanged += comboBox_SelectedIndexChanged;
        }
        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Récupérer la valeur réelle (realName) du ComboBoxEdit
            string realName = cboServers.EditValue?.ToString();

        }

        private void cboServers_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                splashScreenManager1.ShowWaitForm();
                //string realName = cboServers.EditValue?.ToString();
                //MessageBox.Show($"Nom réel sélectionné : {realName}");
                arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(tokenEmail);
                splashScreenManager1.CloseWaitForm();
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue :{m}  : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try
                {
                    splashScreenManager1.CloseWaitForm();
                }
                catch { }
            }
           
        }

        private void tokenEmail_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void checkEdit11_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}