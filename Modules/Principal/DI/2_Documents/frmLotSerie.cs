using arbioApp.Models;
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
    public partial class frmLotSerie : Form
    {
        public frmLotSerie()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (txtLot.Text == "" || string.IsNullOrWhiteSpace(txtLot.Text))
            {
                MessageBox.Show($"Le nom du lot ne peut pas être vide", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else if (dt_peremption.Text == "" || string.IsNullOrWhiteSpace(dt_peremption.Text))
            {
                MessageBox.Show($"La date de péremption ne peut pas être vide", "Erreur", MessageBoxButtons.OK,
                     MessageBoxIcon.Error);
            }
            else
            {
                /*string connectionStringArbio = $"Server=SRV-ARB;" +
                                $"Database=TRANSIT;User ID=Dev;Password=1234;" +
                                $"TrustServerCertificate=True;Connection Timeout=120;";*/

                string connectionStringArbio = $"Server=localhost;" +
                                $"Database=TRANSIT;User ID=Dev;Password=1234;" +
                                $"TrustServerCertificate=True;Connection Timeout=120;";

                using (SqlConnection connection = new SqlConnection(connectionStringArbio))
                {
                    connection.Open();

                    using (SqlTransaction tran = connection.BeginTransaction())
                    {
                        try
                        {
                            // 🔴 Désactiver le trigger
                            using (SqlCommand disableTrig = new SqlCommand(
                                "DISABLE TRIGGER TG_INS_F_LOTSERIE ON F_LOTSERIE",
                                connection, tran))
                            {
                                disableTrig.ExecuteNonQuery();
                            }

                            // ✅ INSERT
                            string sql = @"
                            INSERT INTO F_LOTSERIE
                            (LS_NoSerie, AR_Ref, DE_No, LS_Qte, DL_NoIn,LS_MvtStock,LS_Peremption,LS_QteRestant)
                            VALUES
                            (@LS_NoSerie, @AR_Ref, @DE_No, @Qte, @DL_NoIn,@LS_MvtStock,@LS_Peremption,@LS_QteRestant)";

                            using (SqlCommand cmd = new SqlCommand(sql, connection, tran))
                            {
                                cmd.Parameters.Add("@LS_NoSerie", SqlDbType.VarChar).Value = txtLot.Text;
                                cmd.Parameters.Add("@AR_Ref", SqlDbType.VarChar).Value = txtreference.Text;
                                cmd.Parameters.Add("@DE_No", SqlDbType.Int).Value = int.Parse(textDE_NO.Text);
                                cmd.Parameters.Add("@Qte", SqlDbType.Decimal).Value = decimal.Parse(txtqte.Text);
                                cmd.Parameters.Add("@DL_NoIn", SqlDbType.Int).Value = int.Parse(txtligne.Text);
                                cmd.Parameters.Add("@LS_MvtStock", SqlDbType.Int).Value = 1;
                                cmd.Parameters.Add("@LS_Peremption", SqlDbType.DateTime).Value = dt_peremption.Text;
                                cmd.Parameters.Add("@LS_QteRestant", SqlDbType.Decimal).Value = decimal.Parse(txtqte.Text);

                                cmd.ExecuteNonQuery();
                            }

                            // 🟢 Réactiver le trigger
                            using (SqlCommand enableTrig = new SqlCommand(
                                "ENABLE TRIGGER TG_INS_F_LOTSERIE ON F_LOTSERIE",
                                connection, tran))
                            {
                                enableTrig.ExecuteNonQuery();
                            }

                            tran.Commit();
                        }
                        catch (Exception ex)
                        {
                            tran.Rollback();
                            MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }


                this.Close();
            }
        }

        private void frmLotSerie_Load(object sender, EventArgs e)
        {
            this.ControlBox = false;
        }
    }
}
