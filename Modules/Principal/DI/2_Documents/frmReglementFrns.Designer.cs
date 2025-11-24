namespace arbioApp.Modules.Principal.DI._2_Documents
{
    partial class frmReglementFrns
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
            this.btnCancelRegle = new DevExpress.XtraEditors.SimpleButton();
            this.btnSaveRegle = new DevExpress.XtraEditors.SimpleButton();
            this.btnDeleteRegle = new DevExpress.XtraEditors.SimpleButton();
            this.btnNewRegle = new DevExpress.XtraEditors.SimpleButton();
            this.gcRegle = new DevExpress.XtraGrid.GridControl();
            this.gvRegle = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gcRegle)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvRegle)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.btnCancelRegle);
            this.layoutControl1.Controls.Add(this.btnSaveRegle);
            this.layoutControl1.Controls.Add(this.btnDeleteRegle);
            this.layoutControl1.Controls.Add(this.btnNewRegle);
            this.layoutControl1.Controls.Add(this.gcRegle);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(949, 752, 812, 500);
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1519, 406);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // btnCancelRegle
            // 
            this.btnCancelRegle.Location = new System.Drawing.Point(1420, 367);
            this.btnCancelRegle.Name = "btnCancelRegle";
            this.btnCancelRegle.Size = new System.Drawing.Size(87, 27);
            this.btnCancelRegle.StyleController = this.layoutControl1;
            this.btnCancelRegle.TabIndex = 10;
            this.btnCancelRegle.Text = "Annuler";
            // 
            // btnSaveRegle
            // 
            this.btnSaveRegle.Location = new System.Drawing.Point(1329, 367);
            this.btnSaveRegle.Name = "btnSaveRegle";
            this.btnSaveRegle.Size = new System.Drawing.Size(87, 27);
            this.btnSaveRegle.StyleController = this.layoutControl1;
            this.btnSaveRegle.TabIndex = 9;
            this.btnSaveRegle.Text = "Enregistrer";
            this.btnSaveRegle.Click += new System.EventHandler(this.btnSaveRegle_Click);
            // 
            // btnDeleteRegle
            // 
            this.btnDeleteRegle.Location = new System.Drawing.Point(1410, 84);
            this.btnDeleteRegle.Name = "btnDeleteRegle";
            this.btnDeleteRegle.Size = new System.Drawing.Size(97, 27);
            this.btnDeleteRegle.StyleController = this.layoutControl1;
            this.btnDeleteRegle.TabIndex = 8;
            this.btnDeleteRegle.Text = "Supprimer";
            // 
            // btnNewRegle
            // 
            this.btnNewRegle.Location = new System.Drawing.Point(1410, 53);
            this.btnNewRegle.Name = "btnNewRegle";
            this.btnNewRegle.Size = new System.Drawing.Size(97, 27);
            this.btnNewRegle.StyleController = this.layoutControl1;
            this.btnNewRegle.TabIndex = 7;
            this.btnNewRegle.Text = "Nouveau";
            this.btnNewRegle.Click += new System.EventHandler(this.btnNewRegle_Click);
            // 
            // gcRegle
            // 
            this.gcRegle.Location = new System.Drawing.Point(12, 12);
            this.gcRegle.MainView = this.gvRegle;
            this.gcRegle.Name = "gcRegle";
            this.gcRegle.Size = new System.Drawing.Size(1394, 351);
            this.gcRegle.TabIndex = 6;
            this.gcRegle.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvRegle});
            // 
            // gvRegle
            // 
            this.gvRegle.GridControl = this.gcRegle;
            this.gvRegle.Name = "gvRegle";
            this.gvRegle.OptionsView.ShowFooter = true;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem3,
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.emptySpaceItem1,
            this.emptySpaceItem2,
            this.layoutControlItem4,
            this.layoutControlItem5,
            this.emptySpaceItem3});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1519, 406);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.gcRegle;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(1398, 355);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.btnNewRegle;
            this.layoutControlItem1.Location = new System.Drawing.Point(1398, 41);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(101, 31);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.btnDeleteRegle;
            this.layoutControlItem2.Location = new System.Drawing.Point(1398, 72);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(101, 31);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(1398, 0);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(101, 41);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(1398, 103);
            this.emptySpaceItem2.MaxSize = new System.Drawing.Size(101, 0);
            this.emptySpaceItem2.MinSize = new System.Drawing.Size(101, 10);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(101, 252);
            this.emptySpaceItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.btnSaveRegle;
            this.layoutControlItem4.Location = new System.Drawing.Point(1317, 355);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(91, 31);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(91, 31);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(91, 31);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.btnCancelRegle;
            this.layoutControlItem5.Location = new System.Drawing.Point(1408, 355);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(91, 31);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(91, 31);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(91, 31);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(0, 355);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(1317, 31);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // frmReglementFrns
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1519, 406);
            this.Controls.Add(this.layoutControl1);
            this.Name = "frmReglementFrns";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Saisie des règlement fournisseurs";
            this.Load += new System.EventHandler(this.frmReglementFrns_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gcRegle)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvRegle)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraGrid.GridControl gcRegle;
        private DevExpress.XtraGrid.Views.Grid.GridView gvRegle;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.SimpleButton btnDeleteRegle;
        private DevExpress.XtraEditors.SimpleButton btnNewRegle;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraEditors.SimpleButton btnCancelRegle;
        private DevExpress.XtraEditors.SimpleButton btnSaveRegle;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
    }
}