using arbioApp.Modules.Principal.FINANCE;
using DevExpress.DataAccess.Native.Data;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.Data;
using DevExpress.XtraPivotGrid;
using DevExpress.XtraSpreadsheet;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using System.Windows.Forms;
using System.IO;
using DevExpress.XtraReports.Design;
using System.Reflection;
using DevExpress.XtraTreeList;
using DataTable = System.Data.DataTable;
using DevExpress.XtraSplashScreen;
using System.Data.Entity;
using DevExpress.CodeParser;
using DevExpress.XtraEditors.Repository;
//using Microsoft.Office.Interop.Outlook;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Data.Common;
using DevExpress.XtraEditors.Controls;
using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Models;
using arbioApp.Modules.Principal.DI;


namespace arbioApp.Modules.Principal.RECOUVREMENT
{
    public partial class ucClient : DevExpress.XtraEditors.XtraUserControl
    {
        private static ucClient _instance;
        public static ucClient Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucClient();
                return _instance;
            }
        }
        public ucClient()
        {
            InitializeComponent();

        }
        private SqlConnection connection;
        private DataTable dataTable;
        private SqlDataAdapter dataAdapter;
        private static string connectionString;
        private DataTable serversTable;
        public static int rownum = 0;
        private string server;
        private string database;
        private void ChargerLesServeurs()
        {
            try
            {
                connectionString = $"Server={FrmMdiParent.DataSourceNameValueParent};Database=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True;;Connection Timeout=120;";
                connection = new SqlConnection(connectionString);

                string query = "SELECT * FROM dbo.T_SERVER_CLIENT ORDER BY ADDRESS_IP";

                dataAdapter = new SqlDataAdapter(query, connection);

                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);

                serversTable = new DataTable();
                dataAdapter.Fill(serversTable);
                if (!serversTable.Columns.Contains("IsAlive"))
                    serversTable.Columns.Add("IsAlive");

                foreach (DataRow row in serversTable.Rows)
                {
                    string ip = row["ADDRESS_IP"].ToString();
                    string bdd = row["BDD"].ToString();

                    string testConnStr = $"Server={ip};Database={bdd};User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=3;";
                    bool isAlive = false;

                    try
                    {
                        using (SqlConnection conn = new SqlConnection(testConnStr))
                        {
                            conn.Open();
                            isAlive = true;
                        }
                    }
                    catch
                    {
                        isAlive = false;
                    }

                    row["IsAlive"] = isAlive;
                }



                gridControl2.DataSource = serversTable;
                RepositoryItemCheckEdit checkEdit = new RepositoryItemCheckEdit();
                checkEdit.CheckStyle = CheckStyles.Standard;
                checkEdit.EditValueChanged += CheckEdit_EditValueChanged;

                gridControl2.RepositoryItems.Add(checkEdit);
                gridView2.Columns["actif"].ColumnEdit = checkEdit;

                gridView2.Columns["IsAlive"].OptionsColumn.AllowEdit = false;

                gridView2.BestFitColumns();
                gridView2.RefreshData();
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                LogError($"Échec sur {server} / {database} : {ex.Message}");
            }
        }
        private void CheckEdit_EditValueChanged(object sender, EventArgs e)
        {
            CheckEdit edit = sender as CheckEdit;
            if (edit == null) return;

            GridView view = gridControl2.FocusedView as GridView;
            if (view == null) return;

            int rowHandle = view.FocusedRowHandle;
            if (rowHandle < 0) return;

            // Récupération de la nouvelle valeur booléenne
            bool newValue = Convert.ToBoolean(edit.EditValue);

            // Identification unique de la ligne (ici par ADDRESS_IP + BDD)
            string ip = view.GetRowCellValue(rowHandle, "ADDRESS_IP")?.ToString();
            string bdd = view.GetRowCellValue(rowHandle, "BDD")?.ToString();

            // Mise à jour de la base
            try
            {
                using (SqlConnection conn = new SqlConnection($"Server={FrmMdiParent.DataSourceNameValueParent};Database=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=30;"))
                {
                    string sql = @"UPDATE dbo.T_SERVER_CLIENT 
                           SET actif = @actif 
                           WHERE ADDRESS_IP = @ip AND BDD = @bdd";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@actif", newValue);
                    cmd.Parameters.AddWithValue("@ip", ip);
                    cmd.Parameters.AddWithValue("@bdd", bdd);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                LogError($"Échec sur {server} / {database} : {ex.Message}");
            }
        }


        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            ChargerData();
            splashScreenManager1.CloseWaitForm();
        }
        
        private void ChargerData()
        {
            memoEditLog.Text= "";
            try
            {
                gridControl1.DataSource = null;
                gridView1.Columns.Clear();
                string className = MethodBase.GetCurrentMethod().DeclaringType?.Name ?? "UnknownClass";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = $@"SELECT DB_NAME() DB, CT_Num, CT_Intitule, 
                                    CT_NUM_ARBIOCHEM, CT_INTITULE_ARBIOCHEM
                                    from F_COMPTET WHERE CT_Type = 0";
                    dataTable = new System.Data.DataTable();
                    List<string> connectionStrings = new List<string>();

                    foreach (DataRow row in serversTable.Rows)
                    {
                        if (row["actif"] != DBNull.Value && row["actif"].ToString().ToLower() == "true")
                        {
                            string address = row["ADDRESS_IP"].ToString().Trim();
                            string database = row["BDD"].ToString().Trim();

                            string connStr = $"Server={address};Database={database};User ID=Dev;Password=1234;TrustServerCertificate=True;Connection Timeout=30;";
                            connectionStrings.Add(connStr);
                        }
                    }

                    dataTable = new DataTable(); // Résultat final
                    int rownum = 0;

                    foreach (var connStr in connectionStrings)
                    {
                        var builder = new SqlConnectionStringBuilder(connStr);
                        string server = builder.DataSource;
                        string database = builder.InitialCatalog;

                        try
                        {
                            splashScreenManager1.SetWaitFormDescription($"Connexion à {database}...");

                            using (SqlConnection conn = new SqlConnection(connStr))
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                            {
                                DataTable tempTable = new DataTable();
                                conn.Open();
                                adapter.Fill(tempTable);
                                dataTable.Merge(tempTable);
                                rownum += tempTable.Rows.Count; // Attention, c'est tempTable ici, pas dataTable
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MethodBase m = MethodBase.GetCurrentMethod();
                            LogError($"Échec sur {server} / {database} : {ex.Message}");
                        }
                    }

                    gridControl1.DataSource = dataTable;
                    RepositoryItemHyperLinkEdit hyperlink = new RepositoryItemHyperLinkEdit();
                    hyperlink.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
                    hyperlink.Click += Hyperlink_Click_CT_Num;
                    gridControl1.RepositoryItems.Add(hyperlink);
                    var rowHandle = gridView1.FocusedRowHandle;
                    object value = gridView1.GetRowCellValue(rowHandle, "CT_Num");
                    if (value != null)
                    {
                        gridView1.Columns["CT_NUM_ARBIOCHEM"].ColumnEdit = hyperlink;
                        gridView1.Columns["CT_Num"].ColumnEdit = hyperlink;
                    }
                    
                    

                    gridView1.BestFitColumns();
                }

            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                LogError($"Échec sur {server} / {database} : {ex.Message}");
            }
        }

        private void Hyperlink_Click_CT_Num(object sender, EventArgs e)
        {
            try
            {
                if (splashScreenManager1.IsSplashFormVisible)
                {
                    splashScreenManager1.CloseWaitForm();
                }
                
                    splashScreenManager1.ShowWaitForm();
               
                

                HyperLinkEdit editor = sender as HyperLinkEdit;
                if (editor == null) return;

                GridView view = gridControl1.FocusedView as GridView;
                if (view == null) return;
                Point pt = view.GridControl.PointToClient(Control.MousePosition);
                GridHitInfo hitInfo = view.CalcHitInfo(pt);
                
                if (hitInfo.InRowCell)
                {
                    //string columnName = hitInfo.Column.FieldName; 
                    string columnName = view.FocusedColumn.FieldName;
                    int rowHandle = hitInfo.RowHandle;
                    string cellValue = view.GetRowCellValue(rowHandle, columnName)?.ToString();
                    string db="";
                    string adresseIp = "";
                    if (columnName == "CT_NUM_ARBIOCHEM")
                    {
                        adresseIp = "SRV-ARB";
                        db = "ARBIOCHEM";
                        if (string.IsNullOrEmpty(db)) {
                            splashScreenManager1.CloseWaitForm();
                            return;                           
                        }
                        
                    }
                    else
                    {
                        adresseIp = recupereIp(view.GetRowCellValue(rowHandle, "DB")?.ToString());
                        db = view.GetRowCellValue(rowHandle, "DB")?.ToString(); 
                    }
                        
                    string ctNum = view.GetRowCellValue(rowHandle, columnName)?.ToString();
                        if (string.IsNullOrEmpty(db)) return;

                        string _connectionString = $"Server={adresseIp};" +
                                                        $"Database={db};User ID=Dev;Password=1234;" +
                                                        $"TrustServerCertificate=True;Connection Timeout=120;";

                        string query = $"SELECT CT_Num, CT_Intitule, CT_Classement, " +
                                       $" CT_Adresse, CT_Complement, CT_CodePostal, CT_Ville, CT_Telephone, CT_Site, CT_Encours, CT_ControlEnc, N_Risque, CT_Sommeil " +
                                       $" FROM F_COMPTET WHERE CT_Num = @CTNum";
                    ChargerRisque(_connectionString);

                    // Création d'une liste d'objets anonymes (ou tu peux utiliser une classe dédiée)
                    var data = new List<object>
                        {
                            new { Id = 0, Description = "Contrôle automatique" },
                            new { Id = 1, Description = "Selon code risque" },
                            new { Id = 2, Description = "Compte bloqué" }
                        };

                    // Affectation au LookUpEdit
                    lkTypeCtrl.Properties.DataSource = data;
                    lkTypeCtrl.Properties.DisplayMember = "Description"; // Ce que l'utilisateur voit
                    lkTypeCtrl.Properties.ValueMember = "Id";            // La valeur stockée (0, 1, 2)
                    lkTypeCtrl.Properties.NullText = "Sélectionnez une option"; // Texte par défaut



                    using (SqlConnection conn = new SqlConnection(_connectionString))
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@CTNum", ctNum);
                            conn.Open();

                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    textEditCTNum.Text = reader["CT_Num"].ToString();
                                    textEditCTIntitule.Text = reader["CT_Intitule"].ToString();
                                    textEditCTClassement.Text = reader["CT_Classement"].ToString();
                                    textEditCTAdresse.Text = reader["CT_Adresse"].ToString();
                                    textEditCTComplement.Text = reader["CT_Complement"].ToString();
                                    textEditCTCodePostal.Text = reader["CT_CodePostal"].ToString();
                                    textEditCTVille.Text = reader["CT_Ville"].ToString();
                                    textEditCTTelephone.Text = reader["CT_Telephone"].ToString();
                                    textEditCTSite.Text = reader["CT_Site"].ToString();
                                    tctCtEncours.Text = reader["CT_Encours"].ToString();
                                    lkTypeCtrl.EditValue= Convert.ToInt16(reader["CT_ControlEnc"]);
                                    lkCodeRisque.EditValue = Convert.ToInt16(reader["N_Risque"]);
                                    chkCtSommeil.EditValue = Convert.ToInt16(reader["CT_Sommeil"]);
                            }
                                else
                                {
                                    textEditCTNum.Text = "";
                                    textEditCTIntitule.Text = "";
                                    textEditCTClassement.Text = "";
                                    textEditCTAdresse.Text = "";
                                    textEditCTComplement.Text = "";
                                    textEditCTCodePostal.Text = "";
                                    textEditCTVille.Text = "";
                                    textEditCTTelephone.Text = "";
                                    textEditCTSite.Text = "";
                                }
                            }
                        }
                        string query2 = $"SELECT CT_Num, DO_Date, DO_Piece, AR_Ref, " +
                                       $" DL_Design, DL_Qte, DL_No, DL_PrixUnitaire " +
                                       $" FROM F_DOCLIGNE WHERE CT_Num = @CTNum ORDER BY DO_Date DESC";

                        using (SqlConnection conn = new SqlConnection(_connectionString))
                        using (SqlCommand cmd = new SqlCommand(query2, conn))
                        {
                            cmd.Parameters.AddWithValue("@CTNum", ctNum);
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            DataTable resultTable = new DataTable();
                            conn.Open();
                            adapter.Fill(resultTable);

                            // Étape 4 : Afficher dans gridControl2
                            gridControl3.DataSource = resultTable;
                            gridView3.BestFitColumns();
                        }
                    //}
                }

                
                splashScreenManager1.CloseWaitForm();
            }
            catch (System.Exception ex)
            {
                splashScreenManager1.CloseWaitForm();
                MethodBase m = MethodBase.GetCurrentMethod();
                LogError($"Échec sur {server} / {database} : {ex.Message}");
            }
            
        }
        private string recupereIp(string dbName)
        {
            string addressIp = "";

            for (int i = 0; i < gridView2.RowCount; i++)
            {
                string bdd = gridView2.GetRowCellValue(i, "BDD")?.ToString();

                if (!string.IsNullOrEmpty(bdd) && bdd.Equals(dbName, StringComparison.OrdinalIgnoreCase))
                {
                    addressIp = gridView2.GetRowCellValue(i, "ADDRESS_IP")?.ToString();
                    break;
                }
            }

            return addressIp;
        }

        private void LogError(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new System.Action(() => LogError(message)));
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            memoEditLog.AppendText($"[{timestamp}] [ERREUR] {message}{Environment.NewLine}");
        }

        private void gridView1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.Column.FieldName == "CT_Intitule")
            {
                GridView view = sender as GridView;
                string currentValue = e.CellValue?.ToString();

                // Vérifie si cette valeur existe plus d'une fois
                int count = 0;
                for (int i = 0; i < view.RowCount; i++)
                {
                    object val = view.GetRowCellValue(i, e.Column);
                    if (val != null && val.ToString() == currentValue)
                    {
                        count++;
                        if (count > 1) break;
                    }
                }

                if (count > 1)
                {
                    e.Appearance.BackColor = Color.LightCoral; // ou toute autre couleur
                    e.Appearance.ForeColor = Color.White;
                }
            }
        }

        private void hyperlinkLabelControl2_Click(object sender, EventArgs e)
        {
            gridView1.ShowPrintPreview();
        }
       
        private void ucClient_Load(object sender, EventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            ChargerLesServeurs();
            
            
            splashScreenManager1.CloseWaitForm();
        }
        List<P_CRISQUE> _listeCRisque;
        private void ChargerRisque(string cnx)
        {
            _listeCRisque = GetAllCRisque(cnx);
            lkCodeRisque.Properties.DataSource = _listeCRisque;
            lkCodeRisque.Properties.ValueMember = "cbMarq";
            lkCodeRisque.Properties.DisplayMember = "R_Intitule";
            
        }
        

        private void hyperlinkLabelControl3_Click(object sender, EventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            ChargerLesServeurs();
            splashScreenManager1.CloseWaitForm();
        }

        private void gridView2_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName == "IsAlive")
            {
                bool isAlive = Convert.ToBoolean(e.CellValue);
                Image img = isAlive ? imageCollection1.Images[0] : imageCollection1.Images[1];

                e.Handled = true;
                e.Graphics.DrawImage(img, e.Bounds.X + (e.Bounds.Width - img.Width) / 2, e.Bounds.Y + (e.Bounds.Height - img.Height) / 2);
            }
        }
        private List<P_CRISQUE> GetAllCRisque(string cnx)
        {
            List<P_CRISQUE> crisques = new List<P_CRISQUE>();
            string query = "SELECT cbMarq, R_Intitule FROM P_CRISQUE WHERE R_Intitule <> ''";


            using (SqlConnection conn = new SqlConnection(cnx))
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
                                crisques.Add(new P_CRISQUE
                                {
                                    cbMarq = (int)reader["cbMarq"],
                                    R_Intitule = reader["R_Intitule"].ToString()
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

            return crisques;
        }
    }
}
