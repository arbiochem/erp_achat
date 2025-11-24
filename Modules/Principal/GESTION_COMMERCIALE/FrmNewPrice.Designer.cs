namespace arbioApp.Modules.Principal.GESTION_COMMERCIALE
{
    partial class FrmNewPrice
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmNewPrice));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.txtRef = new DevExpress.XtraEditors.TextEdit();
            this.txtDesign = new DevExpress.XtraEditors.TextEdit();
            this.txtCatTar = new DevExpress.XtraEditors.TextEdit();
            this.txtPrixActuel = new DevExpress.XtraEditors.TextEdit();
            this.btnOkPrice = new DevExpress.XtraEditors.SimpleButton();
            this.btnCancelPrice = new DevExpress.XtraEditors.SimpleButton();
            this.txtNewPrice = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem9 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem10 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem8 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem13 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem12 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtRef.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtDesign.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtCatTar.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPrixActuel.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtNewPrice.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem9)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem10)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem13)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem12)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.gridControl1);
            this.layoutControl1.Controls.Add(this.txtRef);
            this.layoutControl1.Controls.Add(this.txtDesign);
            this.layoutControl1.Controls.Add(this.txtCatTar);
            this.layoutControl1.Controls.Add(this.txtPrixActuel);
            this.layoutControl1.Controls.Add(this.btnOkPrice);
            this.layoutControl1.Controls.Add(this.btnCancelPrice);
            this.layoutControl1.Controls.Add(this.txtNewPrice);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(1004, 66, 812, 500);
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(592, 276);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // gridControl1
            // 
            this.gridControl1.Location = new System.Drawing.Point(12, 202);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(568, 25);
            this.gridControl1.TabIndex = 18;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.Name = "gridView1";
            this.gridView1.OptionsView.ShowGroupPanel = false;
            // 
            // txtRef
            // 
            this.txtRef.Location = new System.Drawing.Point(166, 12);
            this.txtRef.Name = "txtRef";
            this.txtRef.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRef.Properties.Appearance.Options.UseFont = true;
            this.txtRef.Properties.Appearance.Options.UseTextOptions = true;
            this.txtRef.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.txtRef.Properties.ReadOnly = true;
            this.txtRef.Size = new System.Drawing.Size(414, 34);
            this.txtRef.StyleController = this.layoutControl1;
            this.txtRef.TabIndex = 17;
            // 
            // txtDesign
            // 
            this.txtDesign.Location = new System.Drawing.Point(166, 50);
            this.txtDesign.Name = "txtDesign";
            this.txtDesign.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDesign.Properties.Appearance.Options.UseFont = true;
            this.txtDesign.Properties.Appearance.Options.UseTextOptions = true;
            this.txtDesign.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.txtDesign.Properties.ReadOnly = true;
            this.txtDesign.Size = new System.Drawing.Size(414, 34);
            this.txtDesign.StyleController = this.layoutControl1;
            this.txtDesign.TabIndex = 16;
            // 
            // txtCatTar
            // 
            this.txtCatTar.Location = new System.Drawing.Point(166, 88);
            this.txtCatTar.Name = "txtCatTar";
            this.txtCatTar.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCatTar.Properties.Appearance.Options.UseFont = true;
            this.txtCatTar.Properties.Appearance.Options.UseTextOptions = true;
            this.txtCatTar.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.txtCatTar.Properties.ReadOnly = true;
            this.txtCatTar.Size = new System.Drawing.Size(414, 34);
            this.txtCatTar.StyleController = this.layoutControl1;
            this.txtCatTar.TabIndex = 15;
            // 
            // txtPrixActuel
            // 
            this.txtPrixActuel.Location = new System.Drawing.Point(166, 126);
            this.txtPrixActuel.Name = "txtPrixActuel";
            this.txtPrixActuel.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPrixActuel.Properties.Appearance.Options.UseFont = true;
            this.txtPrixActuel.Properties.Appearance.Options.UseTextOptions = true;
            this.txtPrixActuel.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtPrixActuel.Properties.DisplayFormat.FormatString = "n2";
            this.txtPrixActuel.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.txtPrixActuel.Properties.EditFormat.FormatString = "n2";
            this.txtPrixActuel.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.txtPrixActuel.Properties.MaskSettings.Set("MaskManagerType", typeof(DevExpress.Data.Mask.NumericMaskManager));
            this.txtPrixActuel.Properties.MaskSettings.Set("MaskManagerSignature", "allowNull=False");
            this.txtPrixActuel.Properties.MaskSettings.Set("mask", "n2");
            this.txtPrixActuel.Properties.ReadOnly = true;
            this.txtPrixActuel.Size = new System.Drawing.Size(414, 34);
            this.txtPrixActuel.StyleController = this.layoutControl1;
            this.txtPrixActuel.TabIndex = 12;
            // 
            // btnOkPrice
            // 
            this.btnOkPrice.Appearance.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOkPrice.Appearance.Options.UseFont = true;
            this.btnOkPrice.Location = new System.Drawing.Point(364, 231);
            this.btnOkPrice.Name = "btnOkPrice";
            this.btnOkPrice.Size = new System.Drawing.Size(216, 33);
            this.btnOkPrice.StyleController = this.layoutControl1;
            this.btnOkPrice.TabIndex = 14;
            this.btnOkPrice.Text = "Ok";
            //this.btnOkPrice.Click += new System.EventHandler(this.btnOkPrice_Click_1);
            // 
            // btnCancelPrice
            // 
            this.btnCancelPrice.Appearance.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancelPrice.Appearance.Options.UseFont = true;
            this.btnCancelPrice.Location = new System.Drawing.Point(106, 231);
            this.btnCancelPrice.Name = "btnCancelPrice";
            this.btnCancelPrice.Size = new System.Drawing.Size(254, 33);
            this.btnCancelPrice.StyleController = this.layoutControl1;
            this.btnCancelPrice.TabIndex = 13;
            this.btnCancelPrice.Text = "Annuler";
            // 
            // txtNewPrice
            // 
            this.txtNewPrice.Location = new System.Drawing.Point(166, 164);
            this.txtNewPrice.Name = "txtNewPrice";
            this.txtNewPrice.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNewPrice.Properties.Appearance.Options.UseFont = true;
            this.txtNewPrice.Properties.Appearance.Options.UseTextOptions = true;
            this.txtNewPrice.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtNewPrice.Properties.DisplayFormat.FormatString = "n2";
            this.txtNewPrice.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.txtNewPrice.Properties.EditFormat.FormatString = "n2";
            this.txtNewPrice.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.txtNewPrice.Properties.MaskSettings.Set("MaskManagerType", typeof(DevExpress.Data.Mask.NumericMaskManager));
            this.txtNewPrice.Properties.MaskSettings.Set("MaskManagerSignature", "allowNull=False");
            this.txtNewPrice.Properties.MaskSettings.Set("mask", "n2");
            this.txtNewPrice.Size = new System.Drawing.Size(414, 34);
            this.txtNewPrice.StyleController = this.layoutControl1;
            this.txtNewPrice.TabIndex = 11;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem9,
            this.layoutControlItem10,
            this.layoutControlItem8,
            this.layoutControlItem13,
            this.layoutControlItem12,
            this.layoutControlItem6,
            this.layoutControlItem4,
            this.emptySpaceItem1,
            this.layoutControlItem1});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(592, 276);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem9
            // 
            this.layoutControlItem9.Control = this.btnCancelPrice;
            this.layoutControlItem9.Location = new System.Drawing.Point(94, 219);
            this.layoutControlItem9.Name = "layoutControlItem9";
            this.layoutControlItem9.Size = new System.Drawing.Size(258, 37);
            this.layoutControlItem9.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem9.TextVisible = false;
            // 
            // layoutControlItem10
            // 
            this.layoutControlItem10.Control = this.btnOkPrice;
            this.layoutControlItem10.Location = new System.Drawing.Point(352, 219);
            this.layoutControlItem10.Name = "layoutControlItem10";
            this.layoutControlItem10.Size = new System.Drawing.Size(220, 37);
            this.layoutControlItem10.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem10.TextVisible = false;
            // 
            // layoutControlItem8
            // 
            this.layoutControlItem8.Control = this.txtNewPrice;
            this.layoutControlItem8.Location = new System.Drawing.Point(0, 152);
            this.layoutControlItem8.Name = "layoutControlItem8";
            this.layoutControlItem8.Size = new System.Drawing.Size(572, 38);
            this.layoutControlItem8.Text = "NOUVEAU PRIX :";
            this.layoutControlItem8.TextSize = new System.Drawing.Size(142, 16);
            // 
            // layoutControlItem13
            // 
            this.layoutControlItem13.Control = this.txtPrixActuel;
            this.layoutControlItem13.Location = new System.Drawing.Point(0, 114);
            this.layoutControlItem13.Name = "layoutControlItem13";
            this.layoutControlItem13.Size = new System.Drawing.Size(572, 38);
            this.layoutControlItem13.Text = "PRIX ACTUEL :";
            this.layoutControlItem13.TextSize = new System.Drawing.Size(142, 16);
            // 
            // layoutControlItem12
            // 
            this.layoutControlItem12.Control = this.txtCatTar;
            this.layoutControlItem12.Location = new System.Drawing.Point(0, 76);
            this.layoutControlItem12.Name = "layoutControlItem12";
            this.layoutControlItem12.Size = new System.Drawing.Size(572, 38);
            this.layoutControlItem12.Text = "CATEGORIE TARIFAIRE :";
            this.layoutControlItem12.TextSize = new System.Drawing.Size(142, 16);
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.txtDesign;
            this.layoutControlItem6.Location = new System.Drawing.Point(0, 38);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(572, 38);
            this.layoutControlItem6.Text = "DESIGNATION :";
            this.layoutControlItem6.TextSize = new System.Drawing.Size(142, 16);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.txtRef;
            this.layoutControlItem4.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(572, 38);
            this.layoutControlItem4.Text = "REFERENCE :";
            this.layoutControlItem4.TextSize = new System.Drawing.Size(142, 16);
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 219);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(94, 37);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.gridControl1;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 190);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(572, 29);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            this.layoutControlItem1.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            // 
            // FrmNewPrice
            // 
            this.AcceptButton = this.btnOkPrice;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(592, 276);
            this.Controls.Add(this.layoutControl1);
            this.IconOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("FrmNewPrice.IconOptions.SvgImage")));
            this.Name = "FrmNewPrice";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmNewPrice";
            this.Load += new System.EventHandler(this.FrmNewPrice_Load);
            this.Shown += new System.EventHandler(this.FrmNewPrice_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtRef.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtDesign.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtCatTar.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPrixActuel.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtNewPrice.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem9)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem10)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem13)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem12)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.SimpleButton btnOkPrice;
        private DevExpress.XtraEditors.SimpleButton btnCancelPrice;
        private DevExpress.XtraEditors.TextEdit txtNewPrice;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem8;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem9;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem10;
        private DevExpress.XtraEditors.TextEdit txtDesign;
        private DevExpress.XtraEditors.TextEdit txtCatTar;
        private DevExpress.XtraEditors.TextEdit txtPrixActuel;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem13;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem12;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraEditors.TextEdit txtRef;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
    }
}