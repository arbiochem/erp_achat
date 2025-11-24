using arbioApp.Modules.Principal.DI._2_Documents;
using DevExpress.DashboardCommon;
using DevExpress.DashboardWin;
using DevExpress.DataAccess.ConnectionParameters;
using DevExpress.DataAccess.Native;
using DevExpress.DataAccess.Sql;
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

namespace arbioApp.Modules.Principal.Dashboard.CIAL
{
    public partial class ucCial_Copy : DevExpress.XtraEditors.XtraUserControl
    {

        private static ucCial_Copy _instance;
        private System.Data.DataTable dataTable;
        private SqlDataAdapter dataAdapter;
        private SqlConnection connection;
        public static string connectionString;

        public static ucCial_Copy Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucCial_Copy();
                return _instance;
            }
        }
        public ucCial_Copy()
        {
            InitializeComponent();
            DataConnectionParametersBase connParameters = CreateConnectionParameters("MSSqlServer");
            DashboardSqlDataSource sqlDataSource = new DashboardSqlDataSource("SQL Data Source 1", connParameters);
            sqlDataSource.Queries.Add(CreateQuery(""));
            sqlDataSource.Fill();
            dashboardDesigner1.Dashboard = CreateDashboard(sqlDataSource);
            dashboardDesigner1.CreateRibbon();
            initializeFilter();
        }

        void initializeFilter()
        {
            DevExpress.XtraEditors.LookUpEdit cmbDepot = new DevExpress.XtraEditors.LookUpEdit();
            cmbDepot.Dock = DockStyle.Top;
            cmbDepot.Properties.NullText = "Sélectionner un dépôt...";
            cmbDepot.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;

            // 📌 Charger la liste des dépôts depuis la base de données
            cmbDepot.Properties.DataSource = GetDepots();
            cmbDepot.Properties.DisplayMember = "DEPOT";
            cmbDepot.Properties.ValueMember = "DEPOT";

            // 📌 Événement de changement de valeur (appliquer le filtre)
            cmbDepot.EditValueChanged += (s, e) => ApplyFilter(cmbDepot.EditValue?.ToString());

            this.Controls.Add(cmbDepot);  // Ajouter la ComboBox au formulaire
           

        }
        private List<string> GetDepots()
        {
            List<string> depots = new List<string>();
            try
            {
                using (SqlConnection conn = new SqlConnection("Server=localhost;Database=ARBIOCHEM;User Id=DEV;Password=1234;"))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT DISTINCT DEPOT FROM VW_VENTE_JOURNALIER_DEPOT", conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                depots.Add(reader["DEPOT"].ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors du chargement des dépôts : {ex.Message}");
            }
            return depots;
        }
        private void ApplyFilter(string selectedDepot)
        {
            if (dashboardDesigner1.Dashboard == null || string.IsNullOrEmpty(selectedDepot))
                return;

            foreach (var item in dashboardDesigner1.Dashboard.Items)
            {
                if (item is DataDashboardItem dataItem)
                {
                    // Supprimer l'ancien filtre
                    dataItem.FilterString = "";

                    // Appliquer le nouveau filtre
                    dataItem.FilterString = $"[DEPOT] = '{selectedDepot}'";
                }
            }

            dashboardDesigner1.ReloadData(); // Recharger les données après application du filtre
            Button btnRefresh = new Button();
            btnRefresh.Text = "🔄 Rafraîchir";
            btnRefresh.Dock = DockStyle.Top;
            btnRefresh.Click += (s, e) => RefreshData();

            this.Controls.Add(btnRefresh);  // Ajouter le bouton au formulaire
        }

        private void RefreshData()
        {
            try
            {
                dashboardDesigner1.ReloadData();
                Console.WriteLine("✅ Dashboard mis à jour !");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erreur lors du rafraîchissement : {ex.Message}");
            }
        }

        private DataConnectionParametersBase CreateConnectionParameters(string providerName)
        {
            switch (providerName)
            {
                case "MSAccess":
                    return new Access97ConnectionParameters()
                    {
                        FileName = @"Data\nwind.mdb"
                    };
                case "MSSqlServer":
                    return new MsSqlConnectionParameters()
                    {
                        ServerName = "localhost",
                        DatabaseName = "ARBIOCHEM",
                        UserName = "DEV",
                        Password ="1234",
                        AuthorizationType = MsSqlAuthorizationType.SqlServer
                    };
                default:
                    return new XmlFileConnectionParameters()
                    {
                        FileName = @"Data\sales-person.xml"
                    };
            }
        }
        private SqlQuery CreateQuery(string builderName)
        {
            switch (builderName)
            {
                case "fluent":
                    return SelectQueryFluentBuilder
                        .AddTable("VW_VENTE_JOURNALIER_DEPOT")
                        //.SelectColumns("CategoryName", "SalesPerson", "OrderDate", "ExtendedPrice")
                        .Build("Query 1");
                default:
                    return new CustomSqlQuery()
                    {
                        Name = "Query 1",
                        Sql = @"SELECT * FROM VW_VENTE_JOURNALIER_DEPOT"
                    };
            }
        }
        private DevExpress.DashboardCommon.Dashboard CreateDashboard(IDashboardDataSource dataSource)
        {
            DevExpress.DashboardCommon.Dashboard newDashboard = new DevExpress.DashboardCommon.Dashboard();
            newDashboard.DataSources.Add(dataSource);

            ChartDashboardItem chart = new ChartDashboardItem
            {
                DataSource = dataSource,
                DataMember = "Query 1"
            };
            chart.Arguments.Add(new Dimension("LaDate", DateTimeGroupInterval.MonthYear));
            chart.Panes.Add(new ChartPane());
            SimpleSeries salesAmountSeries = new SimpleSeries(SimpleSeriesType.SplineArea)
            {
                Value = new Measure("NetAPayer")
            };
            chart.Panes[0].Series.Add(salesAmountSeries);
            GridDashboardItem grid = new GridDashboardItem
            {
                DataSource = dataSource,
                DataMember = "Query 1"
            };
            grid.Columns.Add(new GridDimensionColumn(new Dimension("DEPOT")));
            grid.Columns.Add(new GridMeasureColumn(new Measure("NetAPayer")));

            newDashboard.Items.AddRange(chart, grid);
            return newDashboard;
        }
    }
}
