namespace arbioApp.Modules.Principal.DI._2_Documents
{
    partial class frmAutorisation
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.hlSave = new DevExpress.XtraEditors.HyperlinkLabelControl();
            this.gcAut = new DevExpress.XtraGrid.GridControl();
            this.gvAut = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colid = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colmailUser = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMenu = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colControl = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colVIEW = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCREATE = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colREAD = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colUPDATE = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDELETE = new DevExpress.XtraGrid.Columns.GridColumn();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gcAut)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvAut)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.hlSave);
            this.layoutControl1.Controls.Add(this.gcAut);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(1108, 250, 812, 500);
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1550, 575);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // hlSave
            // 
            this.hlSave.Location = new System.Drawing.Point(12, 547);
            this.hlSave.Name = "hlSave";
            this.hlSave.Size = new System.Drawing.Size(63, 16);
            this.hlSave.StyleController = this.layoutControl1;
            this.hlSave.TabIndex = 5;
            this.hlSave.Text = "Enregistrer";
            this.hlSave.Click += new System.EventHandler(this.hyperlinkLabelControl1_Click);
            // 
            // gcAut
            // 
            this.gcAut.Location = new System.Drawing.Point(12, 12);
            this.gcAut.MainView = this.gvAut;
            this.gcAut.Name = "gcAut";
            this.gcAut.Size = new System.Drawing.Size(1526, 531);
            this.gcAut.TabIndex = 4;
            this.gcAut.UseEmbeddedNavigator = true;
            this.gcAut.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvAut});
            // 
            // gvAut
            // 
            this.gvAut.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colid,
            this.colmailUser,
            this.colMenu,
            this.colControl,
            this.colVIEW,
            this.colCREATE,
            this.colREAD,
            this.colUPDATE,
            this.colDELETE});
            this.gvAut.GridControl = this.gcAut;
            this.gvAut.Name = "gvAut";
            // 
            // colid
            // 
            this.colid.FieldName = "id";
            this.colid.MinWidth = 25;
            this.colid.Name = "colid";
            this.colid.Visible = true;
            this.colid.VisibleIndex = 0;
            this.colid.Width = 94;
            // 
            // colmailUser
            // 
            this.colmailUser.FieldName = "mailUser";
            this.colmailUser.MinWidth = 25;
            this.colmailUser.Name = "colmailUser";
            this.colmailUser.Visible = true;
            this.colmailUser.VisibleIndex = 1;
            this.colmailUser.Width = 94;
            // 
            // colMenu
            // 
            this.colMenu.FieldName = "Menu";
            this.colMenu.MinWidth = 25;
            this.colMenu.Name = "colMenu";
            this.colMenu.Visible = true;
            this.colMenu.VisibleIndex = 2;
            this.colMenu.Width = 94;
            // 
            // colControl
            // 
            this.colControl.FieldName = "Control";
            this.colControl.MinWidth = 25;
            this.colControl.Name = "colControl";
            this.colControl.Visible = true;
            this.colControl.VisibleIndex = 3;
            this.colControl.Width = 94;
            // 
            // colVIEW
            // 
            this.colVIEW.FieldName = "VIEW";
            this.colVIEW.MinWidth = 25;
            this.colVIEW.Name = "colVIEW";
            this.colVIEW.Visible = true;
            this.colVIEW.VisibleIndex = 4;
            this.colVIEW.Width = 94;
            // 
            // colCREATE
            // 
            this.colCREATE.FieldName = "CREATE";
            this.colCREATE.MinWidth = 25;
            this.colCREATE.Name = "colCREATE";
            this.colCREATE.Visible = true;
            this.colCREATE.VisibleIndex = 5;
            this.colCREATE.Width = 94;
            // 
            // colREAD
            // 
            this.colREAD.FieldName = "READ";
            this.colREAD.MinWidth = 25;
            this.colREAD.Name = "colREAD";
            this.colREAD.Visible = true;
            this.colREAD.VisibleIndex = 6;
            this.colREAD.Width = 94;
            // 
            // colUPDATE
            // 
            this.colUPDATE.FieldName = "UPDATE";
            this.colUPDATE.MinWidth = 25;
            this.colUPDATE.Name = "colUPDATE";
            this.colUPDATE.Visible = true;
            this.colUPDATE.VisibleIndex = 7;
            this.colUPDATE.Width = 94;
            // 
            // colDELETE
            // 
            this.colDELETE.FieldName = "DELETE";
            this.colDELETE.MinWidth = 25;
            this.colDELETE.Name = "colDELETE";
            this.colDELETE.Visible = true;
            this.colDELETE.VisibleIndex = 8;
            this.colDELETE.Width = 94;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1550, 575);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.gcAut;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(1530, 535);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.hlSave;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 535);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(1530, 20);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // frmAutorisation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1550, 575);
            this.Controls.Add(this.layoutControl1);
            this.Name = "frmAutorisation";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmAutorisation";
            this.Load += new System.EventHandler(this.frmAutorisation_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gcAut)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvAut)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraGrid.GridControl gcAut;
        private DevExpress.XtraGrid.Views.Grid.GridView gvAut;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        //private arbappDataSetTableAdapters.DI_AUTORISATIONSTableAdapter dI_AUTORISATIONSTableAdapter;
        private DevExpress.XtraGrid.Columns.GridColumn colid;
        private DevExpress.XtraGrid.Columns.GridColumn colmailUser;
        private DevExpress.XtraGrid.Columns.GridColumn colMenu;
        private DevExpress.XtraGrid.Columns.GridColumn colVIEW;
        private DevExpress.XtraGrid.Columns.GridColumn colCREATE;
        private DevExpress.XtraGrid.Columns.GridColumn colREAD;
        private DevExpress.XtraGrid.Columns.GridColumn colUPDATE;
        private DevExpress.XtraGrid.Columns.GridColumn colDELETE;
        private DevExpress.XtraEditors.HyperlinkLabelControl hlSave;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraGrid.Columns.GridColumn colControl;
    }
}