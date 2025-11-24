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
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using arbioApp.Modules.Helpers;


namespace arbioApp
{
    public partial class FrmProfil : DevExpress.XtraEditors.XtraForm
    {

        public string Username { get; private set; }
        public string UserRole { get; private set; }

        public bool success;
        public string leserveur { get; set; }
        public string labase { get; set; }

        public string pwd
        {
            get; set;
        }
        public FrmProfil()
        {
            InitializeComponent();
            Addserver.LoadSqlServers(comboBoxEdit1servup);
            LoadGroup(leserveur);
            UserRole = FrmMdiParent.UserRole;
            leserveur = FrmMdiParent.DataSourceNameValueParent;
            LoadUser(leserveur);
            Username = FrmMdiParent.Username;
            Addserver.LoadSqlServers(cboServers2);
            cboServers2.EditValueChanged += cboServers2_EditValueChanged;
            if (FrmMdiParent.UserRole != "Administrators")
            {
                simpleButton1.Enabled = false;
            }
        }
        private void cboServers2_EditValueChanged(object sender, EventArgs e)
        {
            btnsubm.Enabled = !string.IsNullOrEmpty(cboServers2.EditValue?.ToString());
        }
        private void FrmConnex_Load(object sender, EventArgs e)
        {
            comboBoxEdit1servup.EditValue = leserveur;
            lknameUp.EditValue = Username;
            lkGroupUp.EditValue = UserRole;
            var permissionsManager = new PermissionsManager($"Data Source={FrmMdiParent.DataSourceNameValueParent};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True");
            string roleName = UserRole; // Exemple : nom de la colonne du rôle
            var permissions = permissionsManager.LoadPermissions(roleName);

            // Appliquer les permissions au formulaire
            permissionsManager.ApplyPermissions(this, permissions, 0);//enabled
        }

        private void LoadGroup(string serverName)
        {
            string query = "SELECT id_group, UserGroup FROM T_GROUP";
            try
            {
                // Connexion à SQL Server
                using (SqlConnection connection = new SqlConnection($"Data Source={FrmMdiParent.DataSourceNameValueParent};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True"))
                {
                    connection.Open();

                    // Créer une commande SQL
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Adapter pour remplir un DataTable
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            // Lier le DataTable au LookUpEdit
                            lkGroupUp.Properties.DataSource = dataTable;
                            lkGroupUp.Properties.DisplayMember = "UserGroup"; // Colonne à afficher
                            lkGroupUp.Properties.ValueMember = "UserGroup";//"id_group";    // Valeur sélectionnée
                            lkGroupUp.Properties.NullText = "Sélectionnez un groupe";

                            lookUpEdit1.Properties.DataSource = dataTable;
                            lookUpEdit1.Properties.DisplayMember = "UserGroup"; // Colonne à afficher
                            lookUpEdit1.Properties.ValueMember = "UserGroup";//"id_group";    // Valeur sélectionnée
                            lookUpEdit1.Properties.NullText = "Sélectionnez un groupe";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Gestion des erreurs
                MessageBox.Show("Erreur lors du chargement des groupes : " + ex.Message);
            }
        }
        private void LoadUser(string serverName)
        {
            string query = "SELECT Username FROM T_UserRole";

            try
            {
                // Connexion à SQL Server
                using (SqlConnection connection = new SqlConnection($"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True"))
                {
                    connection.Open();

                    // Créer une commande SQL
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Adapter pour remplir un DataTable
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dataTableUser = new DataTable();
                            adapter.Fill(dataTableUser);
                            // Lier le DataTable au LookUpEdit
                            lknameUp.Properties.DataSource = dataTableUser;
                            lknameUp.Properties.DisplayMember = "Username"; // Colonne à afficher
                            lknameUp.Properties.ValueMember = "Username";//"id_group";    // Valeur sélectionnée
                            lknameUp.Properties.NullText = "Sélectionnez un utilisateur";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Gestion des erreurs
                MessageBox.Show("Erreur lors du chargement des groupes : " + ex.Message);
            }
        }

        private void btnUpdatePassword_Click(object sender, EventArgs e)
        {
            string username = FrmMdiParent.Username; // L'utilisateur courant connecté.
            string oldPassword = txtOldPassword.Text.Trim();
            string newPassword = txtNewPassword.Text.Trim();
            string confirmPassword = txtConfirmPassword.Text.Trim();
            string serverName = leserveur;
            // Vérification des champs
            if (string.IsNullOrEmpty(oldPassword) || string.IsNullOrEmpty(newPassword) || string.IsNullOrEmpty(confirmPassword))
            {
                MessageBox.Show("Tous les champs sont obligatoires.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (newPassword != confirmPassword)
            {
                MessageBox.Show("Le nouveau mot de passe et la confirmation ne correspondent pas.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!IsPasswordComplexEnough(newPassword))
            {
                MessageBox.Show("Le mot de passe doit contenir au moins 8 caractères, une majuscule, une minuscule, un chiffre et un caractère spécial.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Vérifier que le mot de passe n'est pas vide
            if (string.IsNullOrWhiteSpace(newPassword))
            {
                MessageBox.Show("Veuillez entrer un nouveau mot de passe.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            try
            {
                string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Vérifier l'ancien mot de passe (en le hachant pour la comparaison)
                    string queryCheckPassword = "SELECT Password FROM T_UserRole WHERE Username = @Username";
                    using (SqlCommand command = new SqlCommand(queryCheckPassword, connection))
                    {
                        command.Parameters.AddWithValue("@Username", username);

                        string storedPassword = command.ExecuteScalar()?.ToString();

                        // Comparer l'ancien mot de passe haché avec celui string hashedInputPassword = PasswordHelper.HashPassword(inputPassword);
                        oldPassword = PasswordHelper.HashPassword(oldPassword);
                        if (storedPassword != oldPassword) // Hashage de l'ancien mot de passe pour comparaison
                        {
                            MessageBox.Show("L'ancien mot de passe est incorrect.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    // Mettre à jour avec le nouveau mot de passe (haché)
                    string queryUpdatePassword = "UPDATE T_UserRole SET Password = @NewPassword WHERE Username = @Username";
                    using (SqlCommand command = new SqlCommand(queryUpdatePassword, connection))
                    {
                        command.Parameters.AddWithValue("@NewPassword", PasswordHelper.HashPassword(newPassword)); // Hash du nouveau mot de passe
                        command.Parameters.AddWithValue("@Username", username);

                        int rowsAffected = command.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Le mot de passe a été mis à jour avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ClearPasswordFields();
                        }
                        else
                        {
                            MessageBox.Show("Une erreur s'est produite lors de la mise à jour.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool IsPasswordComplexEnough(string password)
        {
            // Vérifier la longueur minimale
            if (password.Length < 8)
            {
                return false;
            }

            // Vérifier la présence d'une majuscule
            if (!password.Any(char.IsUpper))
            {
                return false;
            }

            // Vérifier la présence d'une minuscule
            if (!password.Any(char.IsLower))
            {
                return false;
            }

            // Vérifier la présence d'un chiffre
            if (!password.Any(char.IsDigit))
            {
                return false;
            }

            // Vérifier la présence d'un caractère spécial
            if (!password.Any(c => !char.IsLetterOrDigit(c)))
            {
                return false;
            }

            return true;
        }

        private void ClearPasswordFields()
        {
            txtOldPassword.Clear();
            txtNewPassword.Clear();
            txtConfirmPassword.Clear();
        }

        private void btnsubm_Click(object sender, EventArgs e)
        {
            // Vérification des champs
            string username = txtregusername.EditValue?.ToString();
            string password = txtregpw.EditValue?.ToString();
            string userGroup = lookUpEdit1.EditValue?.ToString();

            string hashedPassword = PasswordHelper.HashPassword(password);

            if (string.IsNullOrWhiteSpace(username))
            {
                MessageBox.Show("Veuillez entrer un nom d'utilisateur.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Veuillez entrer un mot de passe.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(userGroup))
            {
                MessageBox.Show("Veuillez sélectionner un groupe utilisateur.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Si tous les champs sont remplis, appel de la méthode pour ajouter l'utilisateur
            AjouterUtilisateur(username, hashedPassword, userGroup);
            DateHelper.ApposerCreateDate(serverName, username);

        }

        private void cboServers2_EditValueChanged_1(object sender, EventArgs e)
        {
            serverName = cboServers2.EditValue?.ToString();
            //LoadGroup();
        }
        private void AjouterUtilisateur(string username, string password, string userGroup)
        {
            serverName = cboServers2.EditValue?.ToString();
            string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";

            string query = @"
        INSERT INTO [dbo].[T_UserRole] 
               ([Username], [Password], [UserGroup])
        VALUES (@Username, @Password, @UserGroup)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Ajout des paramètres
                        command.Parameters.AddWithValue("@Username", username);
                        command.Parameters.AddWithValue("@Password", password);
                        command.Parameters.AddWithValue("@UserGroup", userGroup);

                        // Exécution de la commande
                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Utilisateur ajouté avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Aucun enregistrement inséré.", "Échec", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        string serverName;

        private void txtregusername_Validating(object sender, CancelEventArgs e)
        {
            //TextEdit textEdit = sender as TextEdit;
            string email = txtregusername.Text.Trim();

            if (!IsValidEmail(email))
            {
                e.Cancel = true; // Empêche de quitter le contrôle si la validation échoue
                txtregusername.ErrorText = "Veuillez entrer une adresse email valide.";
            }
        }
        private bool IsValidEmail(string email)
        {
            string emailPattern = @"^[^@\s]+@[^@\s]+\.[^@\s]+$"; // Expression régulière pour valider un email
            return Regex.IsMatch(email, emailPattern);
        }

        private void xtraTabGroupes_Click(object sender, EventArgs e)
        {


        }
        private SqlConnection connection;
        private SqlDataAdapter dataAdapter;
        private DataTable dataTable;
        private BindingSource bindingSource;
        private void LoadGroup()
        {

            string connectionString = $"Data Source={leserveur};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";

            connection = new SqlConnection(connectionString);
            connection.Open();

            // Créer un DataTable pour contenir les données
            dataTable = new DataTable();

            // Créer un SqlDataAdapter pour les opérations CRUD
            dataAdapter = new SqlDataAdapter("SELECT UserGroup FROM T_Group WHERE UserGroup <> 'Administrators'", connection);

            // Ajouter un SqlCommandBuilder pour générer les commandes INSERT, UPDATE, DELETE
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);

            // Remplir le DataTable avec les données
            dataAdapter.Fill(dataTable);

            // Créer un BindingSource pour gérer les données
            bindingSource = new BindingSource();
            bindingSource.DataSource = dataTable;

            // Lier le GridControl à la source de données
            gridControl1.DataSource = bindingSource;

            // Activer la barre de navigation intégrée
            gridControl1.UseEmbeddedNavigator = true;
        }

        private void gridView1_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            dataAdapter.Update(dataTable);
        }

        private void gridView1_RowDeleted(object sender, DevExpress.Data.RowDeletedEventArgs e)
        {
            dataAdapter.Update(dataTable);
        }

        private void gridView1_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
        {
            gridView1.SetRowCellValue(e.RowHandle, "UserGroup", "");
        }
        private void SaveChanges()
        {
            dataAdapter.Update(dataTable);
        }

        private SqlDataAdapter dataAdapter2;
        private DataTable dataTable2;
        private BindingSource bindingSource2;
        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if (xtraTabControl1.SelectedTabPage.Name == "xtraTabGroupes")
            {
                LoadGroup();
            }
            else if ((xtraTabControl1.SelectedTabPage.Name == "xtraTabPermissions"))
            {
                LoadPermissions();
            }
        }
        private void LoadPermissions()
        {

            string connectionString = $"Data Source={leserveur};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";

            connection = new SqlConnection(connectionString);
            connection.Open();

            dataTable2 = new DataTable();
            dataAdapter2 = new SqlDataAdapter(@"SELECT 
                                                  *
                                                FROM 
                                                  dbo.T_ComponentPermissions; ", connection);

            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter2);

            dataAdapter2.Fill(dataTable2);

            bindingSource2 = new BindingSource();
            bindingSource2.DataSource = dataTable2;

            gridControl2.DataSource = bindingSource2;

            gridControl2.UseEmbeddedNavigator = true;
            gridView2.OptionsView.ColumnAutoWidth = false;
            gridView2.Columns["ID"].VisibleIndex = -1;
            gridView2.Columns["Module"].VisibleIndex = -1;

            gridView2.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left; // Figer la première colonne
            gridView2.Columns[1].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
        }

        private void gridView2_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            dataAdapter2.Update(dataTable2);
        }

        private void gridView2_RowDeleted(object sender, DevExpress.Data.RowDeletedEventArgs e)
        {
            dataAdapter2.Update(dataTable2);
        }

        private void gridView2_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
        {
            gridView2.SetRowCellValue(e.RowHandle, "ComponentName", "");
        }
        private void SaveChangesPermissions()
        {
            dataAdapter2.Update(dataTable2);
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            UpDateRole(lknameUp.EditValue .ToString() );
        }

        private void UpDateRole(string name)
        {
            string username = name;
            string oldPassword = txtOldPassword.Text.Trim();
            string newPassword = txtNewPassword.Text.Trim();
            string confirmPassword = txtConfirmPassword.Text.Trim();
            string serverName = leserveur;
            

            try
            {
                string connectionString = $"Data Source={serverName};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string queryUpdatePassword = "UPDATE T_UserRole SET UserGroup = @NewUserGroup WHERE Username = @Username";
                    using (SqlCommand command = new SqlCommand(queryUpdatePassword, connection))
                    {
                        command.Parameters.AddWithValue("@NewUserGroup", lkGroupUp.EditValue.ToString());
                        command.Parameters.AddWithValue("@Username", username);

                        int rowsAffected = command.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Une erreur s'est produite lors de la mise à jour.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void lknameUp_EditValueChanged(object sender, EventArgs e)
        {
            string selectedUser = lknameUp.EditValue?.ToString();

            if (!string.IsNullOrEmpty(selectedUser))
            {
                lkGroupUp.EditValue = GetRoleByUser(selectedUser);
            }
        }
        private string GetRoleByUser(string user)
        {
            string role = null;
            string connectionString = $"Data Source={leserveur};Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT UserGroup FROM T_UserRole WHERE User = @User";
                SqlCommand command = new SqlCommand(query, conn);
                command.Parameters.AddWithValue("@User", user);


                object result = command.ExecuteScalar(); // Récupère la première colonne de la première ligne
                if (result != null)
                {
                    role = result.ToString();
                }
            }
            return role;
        }

        private void lknameUp_SelectionChanged(object sender, PopupSelectionChangedEventArgs e)
        {
            
        }
    }

}
