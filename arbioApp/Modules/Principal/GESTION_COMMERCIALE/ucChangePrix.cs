using DevExpress.XtraEditors;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using DataTable = System.Data.DataTable;
using DevExpress.XtraGrid.Views.Grid;
//using Outlook = Microsoft.Office.Interop.Outlook;
using DevExpress.XtraSplashScreen;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using DevExpress.Utils.Drawing.Helpers;
////using Microsoft.Office.Interop.Outlook;
using arbioApp.Modules.Helpers;
using DevExpress.XtraGrid;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraEditors.Repository;
using DevExpress.Spreadsheet;
using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.XtraBars.Navigation;
using System.Net;
using System.Reflection;
using System.Security.Policy;
using Cell = DevExpress.Spreadsheet.Cell;
using DevExpress.PivotGrid.PivotTable;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using Worksheet = DevExpress.Spreadsheet.Worksheet;
using static System.Net.WebRequestMethods;
using System.Net;
using System.Net.Mail;
using DevExpress.SpreadsheetSource;
using MimeKit;
using DevExpress.PivotGrid.OLAP.SchemaEntities;
using DevExpress.XtraLayout.Customization;
using DevExpress.XtraSpreadsheet.DocumentFormats.Xlsb;
using DevExpress.DataAccess.Sql;
using DevExpress.Data.Filtering;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraRichEdit.Model;
using Comment = DevExpress.Spreadsheet.Comment;

namespace arbioApp.Modules.Principal.GESTION_COMMERCIALE
{
    public partial class ucChangePrix : DevExpress.XtraEditors.XtraUserControl
    {

        private static ucChangePrix _instance;
        public static ucChangePrix Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ucChangePrix();
                return _instance;
            }
        }

        private string roleName;
        private void GetPermission()
        {
            roleName = FrmMdiParent.UserRole;

            var permissionsManager = new PermissionsManager($"Data Source={FrmMdiParent.DataSourceNameValueParent};" +
                                                            $"Initial Catalog=arbapp;User ID=Dev;Password=1234;TrustServerCertificate=True");

            var permissions = permissionsManager.LoadPermissions(roleName);

            permissionsManager.ApplyPermissions(this, permissions, 1);

        }

        public ucChangePrix()
        {
            InitializeComponent();

            GetPermission();
            spreadsheetControl1.Document.Worksheets[0].ActiveView.ShowGridlines = false;
            spreadsheetControl1.RemoveShortcutKey(System.Windows.Forms.Keys.S, System.Windows.Forms.Keys.Control);
        }

        string leserveur;
        string UserRole;
        string UserName;
        public static string selectedSite;
        public static string selectedConnex;


        private void envoyerMail(string body)
        {
            //    try
            //    {
            //        Outlook.Application outlookApp = new Outlook.Application();
            //        Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

            //        //if (!string.IsNullOrWhiteSpace(mailItem.To))
            //        //{
            //        mailItem.To = edtTo.EditValue.ToString();
            //        if (!string.IsNullOrWhiteSpace(mailItem.CC))
            //        {
            //            mailItem.CC = edtCc.EditValue.ToString();
            //        }
            //        mailItem.Subject = "Changement de prix";

            //        string htmlBody = "<a>Bonjour,</a> <br />" +
            //                          "<a>Pour modification prix : </a> <br />" +
            //                          "<div>" + body + "</div> <br />" +
            //                          "<a>Cordialement</a>" +
            //                          "<div>" + FrmMdiParent.IDName + "</div> <br />";
            //        mailItem.HTMLBody = htmlBody;


            //        Thread.Sleep(2000);
            //        mailItem.Send();

            //        MessageBox.Show("E-mail envoyé avec succès.");
            //        //}
            //    }
            //    catch (System.Exception ex)
            //    {
            //        MethodBase m = MethodBase.GetCurrentMethod();
            //        MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
        }

        private void ucChangePrix_Load(object sender, EventArgs e)
        {
            arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(edtCc);
            arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(edtTo);
            arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(tokenEditTo);
            arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(tokenEditCc);
            sauvegarderProjetToolStripMenuItem.Enabled = false;

        }
        private void ConfigureValidation()
        {
            Worksheet sheet = spreadsheetControl1.Document.Worksheets[0];

            CellRange range = sheet.Range["G:O"];

            DataValidation validation = sheet.DataValidations.Add(range,
                DataValidationType.Decimal,
                DataValidationOperator.Between,
                0,
                999999.99
            );

            validation.ErrorMessage = "Veuillez saisir uniquement des nombres.";
            validation.ErrorTitle = "Saisie invalide";
            validation.ShowErrorMessage = true;
        }

        private void hyperlinkLabelControl1_Click(object sender, EventArgs e)
        {
            arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(edtTo);
        }

        private void hyperlinkLabelControl2_Click(object sender, EventArgs e)
        {
            arbioApp.Modules.Helpers.HelperToken.LoadEmailsFromDatabase(edtCc);
        }

        private void hyperlinkLabelControl3_Click(object sender, EventArgs e)
        {
            string connectionString = $"Data Source={selectedConnex};Initial Catalog={selectedSite};User ID=Dev;Password=1234;TrustServerCertificate=True";
            //VerifiezNoUpdated(gridControl1, connectionString);
        }

        private void LoadArticle(string site, string db)
        {
            string connectionString;
            if (db.StartsWith("ACTIVO"))
            {

                connectionString = $"Data Source={FrmNewPrix.selectedIp};Initial Catalog={db};User ID=Dev;Password=1234;TrustServerCertificate=True";
            }
            else
            {
                connectionString = $"Data Source={site};Initial Catalog=ARBIOCHEM;User ID=Dev;Password=1234;TrustServerCertificate=True";
            }
            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("dbo.ARB_PivotCategorieTarifaire", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@SITE", db);

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt); // Remplissage du DataTable
                        }
                    }
                    Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
                    spreadsheetControl1.BeginUpdate();

                    worksheet.Import(dt, true, 3, 1);

                    CellRange usedRange = worksheet.GetUsedRange();
                    int firstColumn = worksheet.GetUsedRange().LeftColumnIndex;
                    int firstRow = worksheet.GetUsedRange().TopRowIndex;
                    int lastColumn = worksheet.GetUsedRange().RightColumnIndex;
                    int lastRow = worksheet.GetUsedRange().BottomRowIndex;


                    CellRange headerRange = worksheet.Range.FromLTRB(0, 3, lastColumn, 3);


                    //AJOUTER DATE
                    for (int row = 4; row <= lastRow; row++)
                    {
                        worksheet.Cells[row, 0].Value = DateTime.Now;
                        worksheet.Cells[row, 0].NumberFormat = "dd/MM/yyyy"; // Format date
                    }
                    worksheet.AutoFilter.Apply(headerRange);
                    Cell sourceCell = worksheet.Cells["D4"];

                    //APPLIQUER MISE EN FORME ENTETE
                    headerRange.CopyFrom(sourceCell, PasteSpecial.Formats);
                    for (int row = 3; row <= lastRow; row++)
                    {
                        Cell cell = worksheet.Cells[row, lastColumn];
                        if (cell.Value.IsEmpty)
                        {
                            cell.ClearFormats();
                        }
                    }

                    SetColumnsReadOnly(spreadsheetControl1, new string[] { "DATE", "FAMILLE", "AR_Ref", "DESIGNATION" });
                    SetColumnsReadOnlyWithoutUnderscore(spreadsheetControl1);
                    spreadsheetControl1.EndUpdate();

                }

                catch (System.Exception ex)
                {
                    MethodBase m = MethodBase.GetCurrentMethod();
                    MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void SetColumnsReadOnly(SpreadsheetControl spreadsheet, string[] columnNames)
        {
            try
            {
                IWorkbook workbook = spreadsheet.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                if (worksheet.IsProtected)
                {
                    worksheet.Unprotect("password");
                }

                int headerRow = 3;
                CellRange usedRange = worksheet.GetUsedRange();
                int firstColumn = usedRange.LeftColumnIndex;
                int lastColumn = usedRange.RightColumnIndex;

                // Parcours toutes les colonnes pour vérifier les noms
                for (int col = firstColumn; col <= lastColumn; col++)
                {
                    DevExpress.Spreadsheet.CellValue cellValue = worksheet.Cells[headerRow, col].Value;

                    if (cellValue.IsText)
                    {
                        string columnNameFromSheet = cellValue.TextValue;

                        // Si le nom de la colonne est dans le tableau des colonnes à rendre en lecture seule
                        if (columnNames.Contains(columnNameFromSheet))
                        {
                            worksheet.Columns[col].Protection.Locked = true;
                        }
                        else
                        {
                            worksheet.Columns[col].Protection.Locked = false;
                        }
                    }
                }

                worksheet.Protect("password", WorksheetProtectionPermissions.Default | WorksheetProtectionPermissions.AutoFilters);
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void SetColumnsReadOnlyWithoutUnderscore(SpreadsheetControl spreadsheet)
        {
            try
            {
                IWorkbook workbook = spreadsheet.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                // Déprotéger la feuille si elle est déjà protégée
                if (worksheet.IsProtected)
                {
                    worksheet.Unprotect("password");
                }

                int headerRow = 3; // Ligne des en-têtes (ligne 4 en base 0)
                CellRange usedRange = worksheet.GetUsedRange();
                int firstColumn = usedRange.LeftColumnIndex;
                int lastColumn = usedRange.RightColumnIndex;


                for (int col = 4; col <= lastColumn; col++)
                {
                    DevExpress.Spreadsheet.CellValue cellValue = worksheet.Cells[headerRow, col].Value;

                    if (cellValue.IsText)
                    {
                        string columnNameFromSheet = cellValue.TextValue;

                        // Vérifier si le nom de la colonne ne contient pas de "_"
                        if (!columnNameFromSheet.Contains("_"))
                        {
                            // Verrouiller la colonne (lecture seule)
                            worksheet.Columns[col].Protection.Locked = true;
                        }
                        else
                        {
                            worksheet.Columns[col].Protection.Locked = false;
                        }
                    }
                }

                worksheet.Protect("password", WorksheetProtectionPermissions.Default | WorksheetProtectionPermissions.AutoFilters);
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void SetAllColumnsReadOnly(SpreadsheetControl spreadsheet)
        {
            try
            {
                IWorkbook workbook = spreadsheet.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                // Déprotéger la feuille si elle est déjà protégée
                if (worksheet.IsProtected)
                {
                    worksheet.Unprotect("password");
                }

                int headerRow = 3; // Ligne des en-têtes (ligne 4 en base 0)
                CellRange usedRange = worksheet.GetUsedRange();
                int firstColumn = usedRange.LeftColumnIndex;
                int lastColumn = usedRange.RightColumnIndex;


                for (int col = 0; col <= lastColumn; col++)
                {
                    worksheet.Columns[col].Protection.Locked = true;
                }

                worksheet.Protect("password", WorksheetProtectionPermissions.Default | WorksheetProtectionPermissions.AutoFilters);
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void ouvrirProjetPriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                XtraOpenFileDialog openFileDialog = new XtraOpenFileDialog();
                openFileDialog.InitialDirectory = @"\\Srv-arb\prix$";
                openFileDialog.Filter = "Fichiers projets prix|*.xlsx";
                openFileDialog.Title = "Sélectionner un fichier Excel";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    spreadsheetControl1.LoadDocument(filePath);

                    DevExpress.Spreadsheet.Worksheet sheet = spreadsheetControl1.Document.Worksheets[0];
                    sauvegarderProjetToolStripMenuItem.Enabled = true;

                    SetColumnsReadOnly(spreadsheetControl1, new string[] { "DATE", "FAMILLE", "AR_Ref", "DESIGNATION" });
                    SetColumnsReadOnlyWithoutUnderscore(spreadsheetControl1);

                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void nouveauprixToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmNewPrix frmNewPrix = new FrmNewPrix(_instance);
            frmNewPrix.ShowDialog();
        }
        public void LoadSpreadsheetTemplate(string site, string db, string templatePath)
        {
            spreadsheetControl1.CreateNewDocument();
            spreadsheetControl1.LoadDocument(templatePath);
            LoadArticle(site, db);
            ConfigureValidation();
            sauvegarderProjetToolStripMenuItem.Enabled = true;
        }

        public void sauvegarderProjetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sauvegarder();
            spreadsheetControl1.CreateNewDocument();
            sauvegarderProjetToolStripMenuItem.Enabled = false;
        }

        private string ProjetFileName;
        public void Sauvegarder()
        {
            using (XtraSaveFileDialog saveFileDialog = new XtraSaveFileDialog())
            {
                try
                {
                    saveFileDialog.InitialDirectory = @"\\Srv-arb\prix$";
                    saveFileDialog.Filter = "Fichiers projets prix|*.xlsx";
                    saveFileDialog.DefaultExt = "xlsx";
                    saveFileDialog.AddExtension = true;
                    Cell cell = spreadsheetControl1.Document.Worksheets.ActiveWorksheet.Cells["B1"];
                    object cellValue = cell.Value;
                    saveFileDialog.FileName = cellValue.ToString() + "-" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ProjetFileName = saveFileDialog.FileName;
                        IWorkbook workbook = spreadsheetControl1.Document;
                        Worksheet worksheet = workbook.Worksheets[0];
                        if (worksheet.IsProtected)
                        {
                            worksheet.Unprotect("password");
                        }
                        //worksheet.Protect("password", WorksheetProtectionPermissions.Default | WorksheetProtectionPermissions.AutoFilters);
                        SetAllColumnsReadOnly(spreadsheetControl1);
                        spreadsheetControl1.SaveDocument(saveFileDialog.FileName, DevExpress.Spreadsheet.DocumentFormat.OpenXml);

                        try
                        {
                            if (FrmMdiParent.DataSourceNameValueParent != "(local)")
                            {
                                string fileName = System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                                string fileExtension = System.IO.Path.GetExtension(saveFileDialog.FileName);
                                string hiddenFolderPath = @"\\Srv-arb\log_prix$\";
                                string hiddenFilePath = System.IO.Path.Combine(hiddenFolderPath, fileName + fileExtension);

                                int counter = 1;
                                while (System.IO.File.Exists(hiddenFilePath))
                                {
                                    hiddenFilePath = System.IO.Path.Combine(hiddenFolderPath, $"{fileName}_{counter:D2}{fileExtension}");
                                    counter++;
                                }
                                if (!spreadsheetControl1.Document.Worksheets.ActiveWorksheet.IsProtected)
                                {
                                    spreadsheetControl1.Document.Worksheets.ActiveWorksheet.Protect("password", WorksheetProtectionPermissions.Default);
                                }
                                SetAllColumnsReadOnly(spreadsheetControl1);
                                spreadsheetControl1.SaveDocument(hiddenFilePath, DevExpress.Spreadsheet.DocumentFormat.OpenXml);
                            }

                        }
                        catch (System.Exception ex)
                        {
                            MethodBase m = MethodBase.GetCurrentMethod();
                            MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private Dictionary<(int row, int col), Color> originalColor = new Dictionary<(int, int), Color>();
        private void spreadsheetControl1_CellBeginEdit(object sender, SpreadsheetCellCancelEventArgs e)
        {
            try
            {
                IWorkbook workbook = spreadsheetControl1.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                int row = e.RowIndex;
                int col = e.ColumnIndex;

                DevExpress.Spreadsheet.CellValue cellValue = worksheet.Cells[3, col].Value;

                if (cellValue.IsText)
                {
                    string columnName = cellValue.TextValue;

                    if ((columnName.EndsWith("_DG") && roleName != "DG") ||
                        (columnName.EndsWith("_COM") && roleName != "COMMERCIAL"))
                    {
                        MessageBox.Show("Vous n'avez pas les droits pour modifier cette colonne.", "Accès refusé",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        return;
                    }

                    if (columnName.EndsWith("_DG"))
                    {
                        string baseColumnName = columnName.Replace("_DG", "");

                        for (int refCol = 0; refCol <= worksheet.GetUsedRange().RightColumnIndex; refCol++)
                        {
                            DevExpress.Spreadsheet.CellValue refCellValue = worksheet.Cells[3, refCol].Value;
                            if (refCellValue.IsText && refCellValue.TextValue == baseColumnName)
                            {
                                var key = (row, refCol);
                                if (!originalColor.ContainsKey(key))
                                {
                                    originalColor[key] = worksheet.Cells[row, refCol].Fill.BackgroundColor;
                                }

                                worksheet.Cells[row, refCol].Fill.BackgroundColor = Color.Yellow;
                                break;
                            }
                        }
                    }
                    if (columnName.EndsWith("_COM"))
                    {
                        string baseColumnName = columnName.Replace("_COM", "_DG");

                        for (int refCol = 0; refCol <= worksheet.GetUsedRange().RightColumnIndex; refCol++)
                        {
                            DevExpress.Spreadsheet.CellValue refCellValue = worksheet.Cells[3, refCol].Value;
                            if (refCellValue.IsText && refCellValue.TextValue == baseColumnName)
                            {
                                var key = (row, refCol);
                                if (!originalColor.ContainsKey(key))
                                {
                                    originalColor[key] = worksheet.Cells[row, refCol].Fill.BackgroundColor;
                                }

                                worksheet.Cells[row, refCol].Fill.BackgroundColor = Color.Yellow;
                                break;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private DevExpress.Spreadsheet.CellValue refCellValue;
        private void spreadsheetControl1_CellEndEdit(object sender, SpreadsheetCellValidatingEventArgs e)
        {
            try
            {
                IWorkbook workbook = spreadsheetControl1.Document;
                Worksheet worksheet = workbook.Worksheets[0];
                int row = e.RowIndex;
                int col = e.ColumnIndex;

                DevExpress.Spreadsheet.CellValue cellValue = worksheet.Cells[3, col].Value;

                string baseColumnName;
                if (cellValue.IsText && cellValue.TextValue.EndsWith("_DG"))
                {

                    baseColumnName = cellValue.TextValue.Replace("_DG", "");
                    for (int refCol = 0; refCol <= worksheet.GetUsedRange().RightColumnIndex; refCol++)
                    {
                        DevExpress.Spreadsheet.CellValue refCellValue = worksheet.Cells[3, refCol].Value;
                        if (refCellValue.IsText && refCellValue.TextValue == baseColumnName)
                        {
                            // Récupérer et restaurer la couleur d'origine
                            var key = (row, refCol);
                            if (originalColor.ContainsKey(key))
                            {
                                worksheet.Cells[row, refCol].Fill.BackgroundColor = originalColor[key];
                                originalColor.Remove(key); // Supprimer après restauration
                            }
                            break;
                        }
                    }
                }
                else if (cellValue.IsText && cellValue.TextValue.EndsWith("_COM"))
                {
                    baseColumnName = cellValue.TextValue.Replace("_COM", "");

                    for (int refCol = 0; refCol <= worksheet.GetUsedRange().RightColumnIndex; refCol++)
                    {
                        DevExpress.Spreadsheet.CellValue refCellValue = worksheet.Cells[3, refCol].Value;
                        if (refCellValue.IsText && refCellValue.TextValue == baseColumnName)
                        {
                            // Récupérer et restaurer la couleur d'origine
                            var key = (row, refCol);
                            if (originalColor.ContainsKey(key))
                            {
                                worksheet.Cells[row, refCol].Fill.BackgroundColor = originalColor[key];
                                originalColor.Remove(key); // Supprimer après restauration
                            }
                            break;
                        }
                    }
                }
                else if (cellValue.IsText && cellValue.TextValue.EndsWith("_APP"))
                {
                    baseColumnName = cellValue.TextValue.Replace("_APP", "");

                    for (int refCol = 0; refCol <= worksheet.GetUsedRange().RightColumnIndex; refCol++)
                    {
                        DevExpress.Spreadsheet.CellValue refCellValue = worksheet.Cells[3, refCol].Value;
                        if (refCellValue.IsText && refCellValue.TextValue == baseColumnName)
                        {
                            // Récupérer et restaurer la couleur d'origine
                            var key = (row, refCol);
                            if (originalColor.ContainsKey(key))
                            {
                                worksheet.Cells[row, refCol].Fill.BackgroundColor = originalColor[key];
                                originalColor.Remove(key); // Supprimer après restauration
                            }
                            break;
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
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            SendMail01();
        }


        private string ConvertRangeToHtml(CellRange range)
        {
            DevExpress.Spreadsheet.Worksheet worksheet = spreadsheetControl1.Document.Worksheets.ActiveWorksheet;

            string htmlHeader = $@"<table border='1' style='border-collapse:collapse; text-align:center;'>
                                                        <thead>
                                                            <tr style='background-color: #00A9C0; color: white;'>
                                                                <th rowspan='2'>DATE</th>
                                                                <th rowspan='2'>AR_REF</th>
                                                                <th rowspan='2'>DESIGNATION</th>
                                                                <th colspan='3'>PRIX_ACTUEL</th>
                                                                <th colspan='3'>PRIX_DG</th>
                                                                <th colspan='3'>PRIX_COMMERCIAL</th>
                                                                <th colspan='3'>PRIX_A_APPLIQUER</th>
                                                            </tr>
                                                            <tr style='background-color: #00A9C0; color: white;'>
                                                                <th>PV1</th>
                                                                <th>PV2</th>
                                                                <th>PV3</th>
                                                                <th>PV1</th>
                                                                <th>PV2</th>
                                                                <th>PV3</th>
                                                                <th>PV1</th>
                                                                <th>PV2</th>
                                                                <th>PV3</th>   
                                                                <th>PV1</th>
                                                                <th>PV2</th>
                                                                <th>PV3</th>
                                                            </tr>
                                                        </thead>";


            StringBuilder htmlBuilder = new StringBuilder();
            htmlBuilder.Append(htmlHeader);

            for (int row = 4; row <= range.BottomRowIndex; row++)
            {
                htmlBuilder.Append("<tr>"); // Début de la ligne
                for (int col = range.LeftColumnIndex; col <= range.RightColumnIndex; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    string cellValue = cell.DisplayText; // Texte affiché dans la cellule

                    htmlBuilder.Append($"<td>{WebUtility.HtmlEncode(cellValue)}</td>");
                }
                htmlBuilder.Append("</tr>"); // Fin de la ligne
            }

            htmlBuilder.Append("</table>");
            return htmlBuilder.ToString();
        }

        private void SendMail01()
        {
            var cellTestValue = spreadsheetControl1.ActiveWorksheet.Cells["B1"].Value;
            if (!cellTestValue.IsEmpty)
            {
                if (tokenEditTo.EditValue != null && !string.IsNullOrWhiteSpace(tokenEditTo.EditValue.ToString()))
                {
                    using (XtraSaveFileDialog saveFileDialog = new XtraSaveFileDialog())
                    {
                        try
                        {
                            Sauvegarder();

                            try
                            {
                                //Worksheet worksheet = spreadsheetControl1.Document.Worksheets.ActiveWorksheet;

                                //int startRow = 3;
                                //int lastRow = startRow;

                                //while (!string.IsNullOrEmpty(worksheet.Cells[lastRow, 2].Value.TextValue)) //3 = D
                                //{
                                //    lastRow++;
                                //}

                                //lastRow--;

                                //string rangeString = $"A5:O{lastRow + 1}";
                                //CellRange range = worksheet[rangeString];
                                //string body = ConvertRangeToHtml(range);

                                //if (worksheet.IsProtected)
                                //{
                                //    worksheet.Unprotect("password");
                                //}

                                MailMessage mail = new MailMessage();
                                mail.From = new MailAddress("rija.razanakoto@arbiochem.mg");
                                mail.To.Add(tokenEditTo.EditValue.ToString());
                                var ccAddress = tokenEditCc.EditValue?.ToString();

                                if (!string.IsNullOrEmpty(ccAddress))
                                {
                                    mail.CC.Add(ccAddress);
                                }

                                string htmlBody = "<a>Bonjour,</a> <br />" +
                                                  "<a>Pour modification prix : </a> <br />" +
                                                  "<br />" +
                                                  "<a>Cordialement</a>" +
                                                  "<div>" + FrmMdiParent.IDName + "</div> <br />";
                                mail.Body = htmlBody; // Assigner le corps de l'e-mail
                                mail.IsBodyHtml = true;
                                mail.Subject = "Changement de prix";

                                //MessageBox.Show(body);
                                // Ajouter la pièce jointe
                                Attachment attachment = new Attachment(ProjetFileName);
                                mail.Attachments.Add(attachment);

                                SmtpClient smtpClient = new SmtpClient("smtpauth.moov.mg");
                                smtpClient.Port = 587;
                                smtpClient.Credentials = new NetworkCredential("rija.razanakoto@arbiochem.mg", "LYp@paBIO2400");
                                smtpClient.EnableSsl = true;

                                smtpClient.Send(mail);

                                MessageBox.Show("E-mail envoyé avec succès.");

                                string fileName = System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                                string fileExtension = System.IO.Path.GetExtension(saveFileDialog.FileName);
                                string hiddenFolderPath = @"\\Srv-arb\log_prix$\";
                                string hiddenFilePath = System.IO.Path.Combine(hiddenFolderPath, fileName + fileExtension);

                                int counter = 1;
                                while (System.IO.File.Exists(hiddenFilePath))
                                {
                                    hiddenFilePath = System.IO.Path.Combine(hiddenFolderPath, $"{fileName}_{counter:D2}{fileExtension}");
                                    counter++;
                                }

                            }
                            catch (System.Exception ex)
                            {
                                MethodBase m = MethodBase.GetCurrentMethod();
                                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            spreadsheetControl1.CreateNewDocument();
                        }
                        catch (System.Exception ex)
                        {
                            MethodBase m = MethodBase.GetCurrentMethod();
                            MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Destinataire ?", "To");
                }
            }
            else
            {
                MessageBox.Show("Feuille vide!", "Feuille vide");
            }
        }

        private void SendMail02()
        {
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("rija.razanakoto@arbiochem.mg");
                mail.To.Add("rija.razanakoto@arbiochem.mg");
                mail.Subject = "Test Mail";
                mail.Body = "This is a test email sent from C# application.";

                // Set up the SMTP client
                SmtpClient smtpClient = new SmtpClient("smtpauth.moov.mg");
                smtpClient.Port = 587; // or 25, depending on your SMTP server configuration
                smtpClient.Credentials = new NetworkCredential("rija.razanakoto@arbiochem.mg", "LYp@paBIO2400");
                smtpClient.EnableSsl = true; // Enable SSL if required by your SMTP server

                // Send the email
                smtpClient.Send(mail);
                MessageBox.Show("Email sent successfully!");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}");
            }
        }
        private int GetColumnIndexByValue(Worksheet worksheet, string columnValue)
        {
            // Parcourir la première ligne (ou une autre ligne de référence)
            for (int col = 0; col <= worksheet.GetUsedRange().RightColumnIndex; col++)
            {
                DevExpress.Spreadsheet.CellValue cellValue = worksheet.Cells[3, col].Value; // On suppose ici que les noms de colonne sont sur la première ligne (index 0)

                if (cellValue.IsText && cellValue.TextValue == columnValue)
                {
                    return col; // Retourner l'index de la colonne si la valeur correspond
                }
            }

            // Si la valeur n'est pas trouvée, renvoyer -1 pour indiquer que la colonne n'a pas été trouvée
            return -1;
        }

        private void spreadsheetControl1_CellValueChanged(object sender, SpreadsheetCellEventArgs e)
        {
            try
            {
                Worksheet worksheet = e.Worksheet;
                int row = e.RowIndex;
                int col = e.ColumnIndex;

                Cell modifiedCell = worksheet.Cells[row, col];
                DevExpress.Spreadsheet.CellValue modifiedCellValue = modifiedCell.Value;
                double modifiedValue = modifiedCellValue.NumericValue;

                DevExpress.Spreadsheet.CellValue cellValue = worksheet.Cells[3, col].Value;
                string baseColumnName;
                if (cellValue.IsText && cellValue.TextValue.EndsWith("_DG"))
                {
                    baseColumnName = cellValue.TextValue.Replace("_DG", "");
                    int referenceColumnIndex = GetColumnIndexByValue(worksheet, baseColumnName);
                    Cell referenceCell = worksheet.Cells[row, referenceColumnIndex];
                    DevExpress.Spreadsheet.CellValue referenceCellValue = referenceCell.Value;
                    double referenceValue = referenceCellValue.NumericValue;

                    double percentageDifference = Math.Abs((modifiedValue - referenceValue) / referenceValue) * 100;

                    // Si la différence dépasse 10%, colorier la cellule modifiée en rouge
                    if (percentageDifference > 10)
                    {
                        modifiedCell.Fill.BackgroundColor = System.Drawing.Color.Red;
                        Comment comment;
                        worksheet.ClearComments(modifiedCell);
                        comment = worksheet.Comments.Add(modifiedCell, FrmMdiParent.IDName, (percentageDifference/100).ToString("P1"));
                    }
                    else
                    {
                        modifiedCell.Fill.BackgroundColor = System.Drawing.Color.White;
                        Comment comment;
                        worksheet.ClearComments(modifiedCell);
                    }

                }
                if (cellValue.IsText && cellValue.TextValue.EndsWith("_COM"))
                {
                    baseColumnName = cellValue.TextValue.Replace("_COM", "_DG");
                    int referenceColumnIndex = GetColumnIndexByValue(worksheet, baseColumnName);
                    Cell referenceCell = worksheet.Cells[row, referenceColumnIndex];
                    DevExpress.Spreadsheet.CellValue referenceCellValue = referenceCell.Value;
                    double referenceValue = referenceCellValue.NumericValue;

                    double percentageDifference = Math.Abs((modifiedValue - referenceValue) / referenceValue) * 100;

                    // Si la différence dépasse 10%, colorier la cellule modifiée en rouge
                    if (percentageDifference > 10)
                    {
                        modifiedCell.Fill.BackgroundColor = System.Drawing.Color.Red;
                        Comment comment;
                        worksheet.ClearComments(modifiedCell);
                        comment = worksheet.Comments.Add(modifiedCell, FrmMdiParent.IDName, (percentageDifference/100).ToString("P1"));
                    }
                    else
                    {
                        modifiedCell.Fill.BackgroundColor = System.Drawing.Color.White;
                        Comment comment;
                        worksheet.ClearComments(modifiedCell);
                    }

                }
                if (cellValue.IsText && cellValue.TextValue.EndsWith("_APP"))
                {
                    baseColumnName = cellValue.TextValue.Replace("_COM", "").Replace("_APP","");
                    int referenceColumnIndex = GetColumnIndexByValue(worksheet, baseColumnName);
                    Cell referenceCell = worksheet.Cells[row, referenceColumnIndex];
                    DevExpress.Spreadsheet.CellValue referenceCellValue = referenceCell.Value;
                    double referenceValue = referenceCellValue.NumericValue;

                    double percentageDifference = Math.Abs((modifiedValue - referenceValue) / referenceValue) * 100;

                    // Si la différence dépasse 10%, colorier la cellule modifiée en rouge
                    if (percentageDifference > 10)
                    {
                        modifiedCell.Fill.BackgroundColor = System.Drawing.Color.Red;
                        Comment comment;
                        worksheet.ClearComments(modifiedCell);
                        comment = worksheet.Comments.Add(modifiedCell, FrmMdiParent.IDName, (percentageDifference / 100).ToString("P1"));
                    }
                    else
                    {
                        modifiedCell.Fill.BackgroundColor = System.Drawing.Color.White;
                        Comment comment;
                        worksheet.ClearComments(modifiedCell);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur est survenue : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}


