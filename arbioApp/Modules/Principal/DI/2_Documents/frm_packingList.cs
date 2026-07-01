using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public partial class frm_packingList : DevExpress.XtraEditors.XtraForm
    {
        private static string dbPrincipale = ucDocuments.dbNamePrincipale;
        private static string serveripPrincipale = ucDocuments.serverIpPrincipale;
        //DECLARATIONS
        private static string connectionString = $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";

        private static string connectionString_f = $"Server={serveripPrincipale};Database=TRANSIT;" +
                                                 $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                                 $"Connection Timeout=240;";
        private String _valdopiece;
        private int _id;
        public frm_packingList()
        {
            InitializeComponent();
        }

        public frm_packingList(string val_dopiece)
        {
            this._valdopiece = val_dopiece;
            InitializeComponent();
        }

        public frm_packingList(string val_dopiece,int id)
        {
            this._valdopiece = val_dopiece;
            this._id = id;
            InitializeComponent();
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            string connectionString2 =
                                    $"Server={serveripPrincipale};Database={dbPrincipale};" +
                                    $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
                                    $"Connection Timeout=240;";

            if (string.IsNullOrWhiteSpace(lblval1.Text))
            {
                string insertSql = @"
                    INSERT INTO F_PACKING_LIST
                    (
                        dopiece,
                        type,
                        nombre,
                        eta,
                        etd
                    )
                    VALUES
                    (
                        @do_piece,
                        @type,
                        @nombre,
                        @eta,
                        @etd
                    )";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                using (SqlCommand cmd = new SqlCommand(insertSql, conn))
                {
                    cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                        .Value = _valdopiece.Trim();

                    cmd.Parameters.Add("@type", SqlDbType.VarChar, 20)
                        .Value = cmb_type.Text?.ToString();

                    cmd.Parameters.Add("@nombre", SqlDbType.VarChar)
                        .Value = txtnbr.Text;

                    cmd.Parameters.Add("@eta", SqlDbType.Date)
                        .Value = dteta.Value;

                    cmd.Parameters.Add("@etd", SqlDbType.Date)
                        .Value = dtetd.Value;

                    conn.Open();
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Packing List enregistré avec succès!!!", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                lblval1.Text = "";
                string UPDATESql = @"
                UPDATE F_PACKING_LIST
                    SET type=@type,
                    nombre=@nombre,
                    eta=@eta,
                    etd=@etd
                    WHERE dopiece=@do_piece AND id=@id
                ";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                using (SqlCommand cmd = new SqlCommand(UPDATESql, conn))
                {
                    cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                        .Value =_valdopiece.Trim();
                    cmd.Parameters.Add("@id", SqlDbType.VarChar, 50)
                       .Value = _id;

                    cmd.Parameters.Add("@type", SqlDbType.VarChar, 20)
                        .Value = cmb_type.Text?.ToString();

                    cmd.Parameters.Add("@nombre", SqlDbType.VarChar)
                        .Value = txtnbr.Text;

                    cmd.Parameters.Add("@eta", SqlDbType.Date)
                        .Value = dteta.Value.Date;

                    cmd.Parameters.Add("@etd", SqlDbType.Date)
                        .Value = dtetd.Value.Date;

                    conn.Open();
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Modification Packing List terminée", "Message d'information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            this.Hide();
        }

        private void frm_packingList_Load(object sender, EventArgs e)
        {
            if (_id != null)
            {
                string connectionString2 =
        $"Server={serveripPrincipale};Database={dbPrincipale};" +
        $"User ID=Dev;Password=1234;TrustServerCertificate=True;" +
        $"Connection Timeout=240;";
                string query = "SELECT * FROM F_PACKING_LIST WHERE dopiece = @do_piece AND id=@id";

                using (SqlConnection conn = new SqlConnection(connectionString2))
                {
                    try
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.Add("@do_piece", SqlDbType.VarChar, 50)
                                .Value = _valdopiece.Trim();
                            cmd.Parameters.Add("@id", SqlDbType.VarChar, 50)
                                .Value = _id;

                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                DataTable dt = new DataTable();
                                da.Fill(dt);
                                if (dt.Rows.Count > 0)
                                {
                                    DataRow row = dt.Rows[0];

                                    cmb_type.Text = row[1].ToString();
                                    txtnbr.Text = row[2].ToString();

                                    if (row[3] != DBNull.Value)
                                        dteta.Value = Convert.ToDateTime(row[3]);

                                    if (row[4] != DBNull.Value)
                                        dtetd.Value = Convert.ToDateTime(row[4]);
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
            }
        }
    }
}