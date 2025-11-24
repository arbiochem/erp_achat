using arbioApp.Modules.Principal.DSI;
using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList;
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
using DevExpress.Docs;
using DevExpress.Spreadsheet;

namespace arbioApp.Modules.Principal.FINANCE
{
    public partial class ucFinanceMvtStock : DevExpress.XtraEditors.XtraUserControl
    {
        private static ucFinanceMvtStock _instance;
        public static ucFinanceMvtStock Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucFinanceMvtStock();
                return _instance;
            }
        }

        private string leserveur;
        public ucFinanceMvtStock()
        {
            InitializeComponent();
            leserveur = FrmMdiParent.DataSourceNameValueParent;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            ExecuteQueryForSelectedDatabases();
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            try
            {
                splashScreenManager1.ShowWaitForm();
                Addserver.LoadTreeList(leserveur, treeListDataSource);
                ExpandAllNodes(treeListDataSource);
                treeListDataSource.ExpandAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur est survenue : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (splashScreenManager1.IsSplashFormVisible)
                {
                    splashScreenManager1.CloseWaitForm();
                }
            }

        }
        private void ExpandAllNodes(DevExpress.XtraTreeList.TreeList treeList)
        {
            treeList.BeginUpdate();
            try
            {
                foreach (TreeListNode node in treeList.Nodes)
                {
                    ExpandNodeAndChildren(node);
                }
            }
            finally
            {
                treeList.EndUpdate();
            }
        }
        private void ExpandNodeAndChildren(TreeListNode node)
        {
            node.Expanded = true;
            foreach (TreeListNode childNode in node.Nodes)
            {
                ExpandNodeAndChildren(childNode);
            }
        }
        private void ExecuteQueryForSelectedDatabases()
        {
            // 1. Récupérer les dates sélectionnées
            DateTime dateDebut = dateEditD.DateTime;
            DateTime dateFin = dateEditF.DateTime;

            // 2. Préparer un DataTable pour combiner les résultats
            DataTable combinedResults = new DataTable();

            // 3. Itérer à travers les nœuds parents (serveurs) dans le TreeList
            foreach (TreeListNode serverNode in treeListDataSource.Nodes)
            {
                if (serverNode.CheckState == CheckState.Checked || serverNode.Checked) // Si le serveur ou une base est cochée
                {
                    string serverName = serverNode.GetValue(1).ToString();

                    // 4. Itérer à travers les nœuds enfants (bases de données)
                    foreach (TreeListNode dbNode in serverNode.Nodes)
                    {
                        if (dbNode.CheckState == CheckState.Checked) // Si la base de données est cochée
                        {
                            string databaseName = dbNode.GetValue(0).ToString();

                            // Construire une chaîne de connexion pour cette base de données
                            string connectionString = $"Data Source={serverName};Initial Catalog={databaseName};User ID=Dev;Password=1234;TrustServerCertificate=True";

                            try
                            {
                                splashScreenManager1.ShowWaitForm();
                                using (SqlConnection connection = new SqlConnection(connectionString))
                                {
                                    connection.Open();

                                    // 5. Exécuter la requête pour chaque base
                                    string query = @"SELECT DB_NAME() AS SOCIETE,
	                        *
                        FROM (
	                        SELECT
                                DATEPART(yy, pvt.DO_Date) AS ANNEE, 
                                pvt.DO_Date AS DATE_DU,
                                pvt.DL_No,
                                CONVERT(VARCHAR(8), CAST(@DATEDEBUT AS DATE), 112) + '-' + CONVERT(VARCHAR(8), CAST(@DATEFIN AS DATE), 112) AS PERIODE,
                                CASE LEN(DATEPART(WW, pvt.DO_Date))
                                WHEN 1
        	                        THEN 'S0' + CONVERT(VARCHAR(1),DATEPART(WW, pvt.DO_Date))
                                WHEN 2
        	                        THEN 'S' + CONVERT(VARCHAR(2),DATEPART(WW, pvt.DO_Date))
                                END AS SEMAINE_DE,
		                        pvt.REFERENCE AS CODE_ARTICLE,
		                        pvt.PIECE,
                                pvt.DO_Ref,
		                        pvt.DESIGNATION,
                                --pvt.DL_Qte,
		                        ISNULL(pvt.STOCK_INITIAL,0) AS STOCK_INITIAL,
		                        ISNULL(pvt.VENTE,0) AS VENTE,
		                        ISNULL(pvt.MS,0) AS MS,
		                        ISNULL(pvt.ACHAT,0) AS ACHAT,
		                        ISNULL(pvt.ME,0) AS ME,
                                CASE
                                WHEN pvt.MOUVEMENT = 'Sortie'
                                THEN - ISNULL(pvt.MT,0)
                                ELSE ISNULL(pvt.MT,0)
                                END AS MT,
		                        ISNULL(pvt.AUTRES,0) AS AUTRES,
                                --pvt.DL_MvtStock,
                                --pvt.MOUVEMENT,
                                pvt.DL_PrixUnitaire AS PU,
		                        pvt.DL_MontantHT AS MontantHT,
		                        --pvt.STD_DENO,
		                        ISNULL(pvt.STOCK_FINAL,0) AS STOCK_FINAL,
                                CASE
                                WHEN ISNULL(pvt.VENTE,0) <> 0 AND ISNULL(pvt.DL_MontantHT,0) <> 0
                                THEN (((pvt.DL_MontantHT - pvt.DL_CMUP) * ISNULL(pvt.VENTE,0) * 1) / ISNULL(pvt.DL_MontantHT,0) / ISNULL(pvt.VENTE,0))
                                ELSE 0
                                END AS  [COHERENCE DE MARGE],
                                pvt.DEPOT,
		                        pvt.CT_Num,
		                        pvt.CT_Intitule AS TIERS,
                                pvt.OBSERVATION
	                        FROM (
		                        SELECT dbo.F_DOCLIGNE.DO_Date,
			                        A.REFERENCE,
			                        dbo.F_DOCLIGNE.DL_Design AS DESIGNATION,
			                        dbo.F_DOCLIGNE.DO_Piece AS PIECE,
                                    dbo.F_DOCLIGNE.DO_Ref,
                                    dbo.F_DOCLIGNE.DL_No,
			                        A.STOCK_INITIAL,
			                        A.STOCK_FINAL,
			                        CASE 
				                        WHEN dbo.F_DOCLIGNE.DO_Type BETWEEN 0
						                        AND 8
					                        THEN 'VENTE'
				                        WHEN dbo.F_DOCLIGNE.DO_Type = 21
					                        THEN 'MS'
				                        WHEN dbo.F_DOCLIGNE.DO_Type BETWEEN 10
						                        AND 18
					                        THEN 'ACHAT'
				                        WHEN dbo.F_DOCLIGNE.DO_Type = 20
					                        THEN 'ME'
				                        WHEN dbo.F_DOCLIGNE.DO_Type = 23
					                        THEN 'MT'
				                        ELSE 'AUTRES'
				                        END AS TYPE,
			                        dbo.F_DOCLIGNE.DL_Qte,
                                    dbo.F_DOCLIGNE.DL_PrixUnitaire,
			                        dbo.F_DOCLIGNE.DL_MontantHT,
                                    dbo.F_DOCLIGNE.DL_MvtStock,
			                        CASE
                                    WHEN dbo.F_DOCLIGNE.DL_MvtStock = 1
                                    THEN 'Entrée'
                                    WHEN dbo.F_DOCLIGNE.DL_MvtStock = 3
                                    THEN 'Sortie'
                                    ELSE ''
                                    END AS MOUVEMENT,
			                        A.STD_DENO,
			                        A.DEPOT,
			                        dbo.F_DOCLIGNE.CT_Num,
			                        dbo.F_COMPTET.CT_Intitule,
                                    dbo.F_DOCLIGNE.DL_CMUP,
			                        '' AS OBSERVATION
		                        FROM dbo.F_DOCLIGNE
		                        INNER JOIN (
			                        SELECT dbo.DP_STOCK_A_DATE.STD_ARREF AS REFERENCE,
				                        dbo.DP_STOCK_A_DATE.STD_DENO,
				                        dbo.DP_STOCK_A_DATE.STD_ARDESING,
				                        dbo.DP_STOCK_A_DATE.STD_DEINTITULE AS DEPOT,
				                        SUM(CASE 
						                        WHEN dbo.DP_STOCK_A_DATE.STD_DLDATEBL <= convert(VARCHAR, DATEADD(day, - 1, @DATEDEBUT), 103)
							                        THEN dbo.DP_STOCK_A_DATE.STD_QTE
                                                    ELSE 0
						                        END) AS STOCK_INITIAL,
				                        SUM(CASE 
						                        WHEN dbo.DP_STOCK_A_DATE.STD_DLDATEBL <= convert(VARCHAR, DATEADD(day, - 0, @DATEFIN), 103)
							                        THEN dbo.DP_STOCK_A_DATE.STD_QTE
                                                    ELSE 0
						                        END) AS STOCK_FINAL
			                        FROM dbo.DP_STOCK_A_DATE
			                        GROUP BY dbo.DP_STOCK_A_DATE.STD_ARREF,
				                        dbo.DP_STOCK_A_DATE.STD_DENO,
				                        dbo.DP_STOCK_A_DATE.STD_ARDESING,
				                        dbo.DP_STOCK_A_DATE.STD_DEINTITULE
			                        ) A ON (dbo.F_DOCLIGNE.AR_Ref = A.REFERENCE)
			                        AND (dbo.F_DOCLIGNE.DE_No = A.STD_DENO)
		                        LEFT OUTER JOIN dbo.F_COMPTET ON (dbo.F_DOCLIGNE.CT_Num = dbo.F_COMPTET.CT_Num)
		                        WHERE (dbo.F_DOCLIGNE.DL_MvtStock = 1 OR dbo.F_DOCLIGNE.DL_MvtStock = 3)) up
	                        PIVOT(SUM(up.DL_Qte) FOR TYPE IN (
				                        VENTE,
				                        MS,
				                        ACHAT,
				                        ME,
				                        MT,
				                        AUTRES
				                        )) AS pvt
	                        ) H
                        WHERE (H.DATE_DU BETWEEN @DATEDEBUT
		                        AND @DATEFIN)";
                                    using (SqlCommand command = new SqlCommand(query, connection))
                                    {
                                        command.Parameters.AddWithValue("@DATEDEBUT", dateDebut);
                                        command.Parameters.AddWithValue("@DATEFIN", dateFin);
                                        command.CommandTimeout = 0;
                                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                                        {
                                            DataTable tempTable = new DataTable();
                                            adapter.Fill(tempTable);

                                            //// Ajouter une colonne pour identifier la base de données et le serveur source
                                            //if (!tempTable.Columns.Contains("SourceServer"))
                                            //    tempTable.Columns.Add("SourceServer", typeof(string));
                                            //if (!tempTable.Columns.Contains("SourceDatabase"))
                                            //    tempTable.Columns.Add("SourceDatabase", typeof(string));

                                            //foreach (DataRow row in tempTable.Rows)
                                            //{
                                            //    row["SourceServer"] = serverName;
                                            //    row["SourceDatabase"] = databaseName;
                                            //}

                                            // Combiner les résultats
                                            if (combinedResults.Columns.Count == 0)
                                            {
                                                combinedResults = tempTable.Clone(); // Copier la structure
                                            }
                                            combinedResults.Merge(tempTable); // Fusionner les données
                                        }
                                    }
                                }
                                splashScreenManager1.CloseWaitForm();
                            }
                            catch (Exception ex)
                            {
                                splashScreenManager1.CloseWaitForm();
                                MessageBox.Show($"Erreur pour la base '{databaseName}' sur le serveur '{serverName}' : {ex.Message}");
                            }
                        }
                    }
                }
            }

            // 6. Afficher les résultats dans un GridControl
            //gridControlResults.DataSource = combinedResults;
            spreadsheetControl1.Document.Worksheets[0].Import(combinedResults, true, 0, 0);
        }

        private void treeListDataSource_BeforeCheckNode(object sender, CheckNodeEventArgs e)
        {
            TreeListNode node = e.Node;
            if (node.Checked)
            {
                node.UncheckAll();
            }
            else
            {
                node.CheckAll();
            }
            while (node.ParentNode != null)
            {
                node = node.ParentNode;
                bool oneOfChildIsChecked = OneOfChildsIsChecked(node);
                if (oneOfChildIsChecked)
                {
                    node.CheckState = CheckState.Checked;
                }
                else
                {
                    node.CheckState = CheckState.Unchecked;
                }
            }
        }
        private bool OneOfChildsIsChecked(TreeListNode node)
        {
            bool result = false;
            foreach (TreeListNode item in node.Nodes)
            {
                if (item.CheckState == CheckState.Checked)
                {
                    result = true;
                }
            }
            return result;
        }

        private void ucFinanceMvtStock_Load(object sender, EventArgs e)
        {
            string leserveur = FrmMdiParent.DataSourceNameValueParent;
            string UserRole = FrmMdiParent.UserRole;
        }

        private void spreadsheetCommandBarButtonItem9_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }
    }
}
