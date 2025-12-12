namespace arbioApp.Modules.Principal.DI._2_Documents
{
    partial class frmAutorisations_achat
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
            this.dataLayoutControl1 = new DevExpress.XtraDataLayout.DataLayoutControl();
            this.hyperlinkLabelControl1 = new DevExpress.XtraEditors.HyperlinkLabelControl();
            this.gd_autorisation_achat = new DevExpress.XtraGrid.GridControl();
            this.gv_achat = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.Colonne_Id = new DevExpress.XtraGrid.Columns.GridColumn();
            this.ColsmailUser = new DevExpress.XtraGrid.Columns.GridColumn();
            this.type_document = new DevExpress.XtraGrid.Columns.GridColumn();
            this.view = new DevExpress.XtraGrid.Columns.GridColumn();
            this.create = new DevExpress.XtraGrid.Columns.GridColumn();
            this.update = new DevExpress.XtraGrid.Columns.GridColumn();
            this.delete = new DevExpress.XtraGrid.Columns.GridColumn();
            this.transform = new DevExpress.XtraGrid.Columns.GridColumn();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataLayoutControl1)).BeginInit();
            this.dataLayoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gd_autorisation_achat)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gv_achat)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            this.SuspendLayout();
            // 
            // dataLayoutControl1
            // 
            this.dataLayoutControl1.Controls.Add(this.hyperlinkLabelControl1);
            this.dataLayoutControl1.Controls.Add(this.gd_autorisation_achat);
            this.dataLayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataLayoutControl1.Location = new System.Drawing.Point(0, 0);
            this.dataLayoutControl1.Name = "dataLayoutControl1";
            this.dataLayoutControl1.Root = this.layoutControlGroup1;
            this.dataLayoutControl1.Size = new System.Drawing.Size(1211, 545);
            this.dataLayoutControl1.TabIndex = 0;
            this.dataLayoutControl1.Text = "dataLayoutControl1";
            // 
            // hyperlinkLabelControl1
            // 
            this.hyperlinkLabelControl1.Location = new System.Drawing.Point(12, 517);
            this.hyperlinkLabelControl1.Name = "hyperlinkLabelControl1";
            this.hyperlinkLabelControl1.Size = new System.Drawing.Size(63, 16);
            this.hyperlinkLabelControl1.StyleController = this.dataLayoutControl1;
            this.hyperlinkLabelControl1.TabIndex = 5;
            this.hyperlinkLabelControl1.Text = "Enregistrer";
            this.hyperlinkLabelControl1.Click += new System.EventHandler(this.hyperlinkLabelControl1_Click);
            // 
            // gd_autorisation_achat
            // 
            this.gd_autorisation_achat.Location = new System.Drawing.Point(12, 12);
            this.gd_autorisation_achat.MainView = this.gv_achat;
            this.gd_autorisation_achat.Name = "gd_autorisation_achat";
            this.gd_autorisation_achat.Size = new System.Drawing.Size(1187, 501);
            this.gd_autorisation_achat.TabIndex = 4;
            this.gd_autorisation_achat.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gv_achat});
            // 
            // gv_achat
            // 
            this.gv_achat.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.Colonne_Id,
            this.ColsmailUser,
            this.type_document,
            this.view,
            this.create,
            this.update,
            this.delete,
            this.transform});
            this.gv_achat.GridControl = this.gd_autorisation_achat;
            this.gv_achat.Name = "gv_achat";
            // 
            // Colonne_Id
            // 
            this.Colonne_Id.Caption = "Id";
            this.Colonne_Id.MinWidth = 25;
            this.Colonne_Id.Name = "Colonne_Id";
            this.Colonne_Id.Visible = true;
            this.Colonne_Id.VisibleIndex = 0;
            this.Colonne_Id.Width = 94;
            // 
            // ColsmailUser
            // 
            this.ColsmailUser.Caption = "Mail User";
            this.ColsmailUser.FieldName = "mailUser";
            this.ColsmailUser.MinWidth = 25;
            this.ColsmailUser.Name = "ColsmailUser";
            this.ColsmailUser.Visible = true;
            this.ColsmailUser.VisibleIndex = 1;
            this.ColsmailUser.Width = 94;
            // 
            // type_document
            // 
            this.type_document.Caption = "Type de document";
            this.type_document.MinWidth = 25;
            this.type_document.Name = "type_document";
            this.type_document.Visible = true;
            this.type_document.VisibleIndex = 2;
            this.type_document.Width = 94;
            // 
            // view
            // 
            this.view.Caption = "VIEW";
            this.view.MinWidth = 25;
            this.view.Name = "view";
            this.view.Visible = true;
            this.view.VisibleIndex = 3;
            this.view.Width = 94;
            // 
            // create
            // 
            this.create.Caption = "CREATE";
            this.create.MinWidth = 25;
            this.create.Name = "create";
            this.create.Visible = true;
            this.create.VisibleIndex = 4;
            this.create.Width = 94;
            // 
            // update
            // 
            this.update.Caption = "UPDATE";
            this.update.MinWidth = 25;
            this.update.Name = "update";
            this.update.Visible = true;
            this.update.VisibleIndex = 5;
            this.update.Width = 94;
            // 
            // delete
            // 
            this.delete.Caption = "DELETE";
            this.delete.MinWidth = 25;
            this.delete.Name = "delete";
            this.delete.Visible = true;
            this.delete.VisibleIndex = 6;
            this.delete.Width = 94;
            // 
            // transform
            // 
            this.transform.Caption = "TRANSFORM";
            this.transform.MinWidth = 25;
            this.transform.Name = "transform";
            this.transform.Visible = true;
            this.transform.VisibleIndex = 7;
            this.transform.Width = 94;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem3,
            this.layoutControlItem4});
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(1211, 545);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.gd_autorisation_achat;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(1191, 505);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.hyperlinkLabelControl1;
            this.layoutControlItem4.Location = new System.Drawing.Point(0, 505);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(1191, 20);
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // frmAutorisations_achat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1211, 545);
            this.Controls.Add(this.dataLayoutControl1);
            this.Name = "frmAutorisations_achat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Autorisations type de documents";
            this.Activated += new System.EventHandler(this.frmAutorisations_achat_Activated);
            this.Load += new System.EventHandler(this.frmAutorisations_achat_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataLayoutControl1)).EndInit();
            this.dataLayoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gd_autorisation_achat)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gv_achat)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
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
        private DevExpress.XtraDataLayout.DataLayoutControl dataLayoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraGrid.GridControl gd_autorisation_achat;
        private DevExpress.XtraGrid.Views.Grid.GridView gv_achat;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.HyperlinkLabelControl hyperlinkLabelControl1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraGrid.Columns.GridColumn Colonne_Id;
        private DevExpress.XtraGrid.Columns.GridColumn ColsmailUser;
        private DevExpress.XtraGrid.Columns.GridColumn type_document;
        private DevExpress.XtraGrid.Columns.GridColumn view;
        private DevExpress.XtraGrid.Columns.GridColumn create;
        private DevExpress.XtraGrid.Columns.GridColumn update;
        private DevExpress.XtraGrid.Columns.GridColumn delete;
        private DevExpress.XtraGrid.Columns.GridColumn transform;
    }
}