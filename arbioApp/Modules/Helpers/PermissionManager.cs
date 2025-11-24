using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Sql;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using System.Windows.Forms;
using DevExpress.XtraTab;

namespace arbioApp
{

    public class PermissionsManager
    {
        private readonly string _connectionString;

        public PermissionsManager(string connectionString)
        {
            _connectionString = connectionString;
        }

        // Charger les permissions depuis la base de données
        public Dictionary<string, bool> LoadPermissions(string roleName)
        {
            var components = new Dictionary<string, bool>();

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                string query = @"
                SELECT ComponentName, [" + roleName + @"] 
                FROM dbo.T_ComponentPermissions";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string componentName = reader["ComponentName"].ToString();
                            bool isEnabled = reader.GetBoolean(1); // La deuxième colonne est la valeur du rôle
                            components[componentName] = isEnabled;
                        }
                    }
                }
            }

            return components;
        }

        
        public void ApplyPermissions(Control container, Dictionary<string, bool> permissions, int visibilityOption)
        {
            // Parcourir tous les contrôles dans le conteneur
            foreach (Control component in container.Controls)
            {


                // Vérifier si le composant existe dans la liste des permissions
                if (permissions.ContainsKey(component.Name))
                {
                    if (visibilityOption == 0)
                    {
                        component.Enabled = permissions[component.Name]; // Applique Enabled
                    }
                    else if (visibilityOption == 1)
                    {
                        component.Visible = permissions[component.Name]; // Applique Visible
                    }
                }

                if (component.HasChildren)
                {
                    ApplyPermissions(component, permissions, visibilityOption);
                }


            }

            // Si le conteneur est un AccordionControl, parcourir ses éléments
            if (container is DevExpress.XtraBars.Navigation.AccordionControl accordionControl)
            {
                accordionControl.ForEachElement((el) =>
                {
                    if (permissions.ContainsKey(el.Name))
                    {
                        if (visibilityOption == 0)
                        {
                            el.Enabled = permissions[el.Name]; // Applique Enabled
                        }
                        else if (visibilityOption == 1)
                        {
                            el.Visible = permissions[el.Name]; // Applique Visible
                        }
                    }
                });
            }
            if (container is DevExpress.XtraTab.XtraTabControl tabControl)
            {
                foreach (DevExpress.XtraTab.XtraTabPage tabPage in tabControl.TabPages)
                {
                    // Vérifier si le nom de l'onglet figure dans les permissions
                    if (permissions.ContainsKey(tabPage.Name))
                    {
                        // Appliquer la permission selon le choix
                        if (visibilityOption == 1)
                        {
                            tabPage.PageVisible = permissions[tabPage.Name]; // Applique Enabled
                        }
                        else if (visibilityOption == 0)
                        {
                            tabPage.PageVisible = permissions[tabPage.Name]; // Applique Visible
                        }
                    }
                }
            }
            if (container is DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm fluentForm)
            {
               // Accéder à FluentDesignFormControl via le FluentForm
               var fluentControl = fluentForm.FluentDesignFormControl;
                
                if (fluentControl.Manager != null)
                {
                    foreach (var item in fluentControl.Manager.Items)
                    {
                        if (item is DevExpress.XtraBars.BarButtonItem barButtonItem)
                        {
                            if (permissions.ContainsKey(barButtonItem.Name))
                            {
                                if (visibilityOption == 0)
                                {
                                    barButtonItem.Enabled = permissions[barButtonItem.Name]; // Applique Enabled
                                }
                                else if (visibilityOption == 1)
                                {
                                    barButtonItem.Visibility = permissions[barButtonItem.Name]
                                        ? DevExpress.XtraBars.BarItemVisibility.Always
                                        : DevExpress.XtraBars.BarItemVisibility.Never; // Applique Visible
                                }
                            }
                        }
                    }
                }
            }
            
        }


    }
}
