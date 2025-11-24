namespace arbioApp
{
    partial class FrmMdiParent
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMdiParent));
            this.popupMenu2 = new DevExpress.XtraBars.PopupMenu(this.components);
            this.ribbonPage2 = new DevExpress.XtraBars.Ribbon.RibbonPage();
            this.splashScreenManager1 = new DevExpress.XtraSplashScreen.SplashScreenManager(this, typeof(global::arbioApp.Modules.Principal.WaitForm1), true, true);
            this.Container = new DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormContainer();
            this.accordionControl1 = new DevExpress.XtraBars.Navigation.AccordionControl();
            this.mnuDSI = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuVReleaseUser = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuUpDatePrix = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuDRH = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuFINANCE = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuMvgStk = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.accordionControlElement2 = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuGESCOM = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuEditPrix = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuImportExport = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.accordionControlElement3 = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuDashboard = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuDashCIAL = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuDashCPTA = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuDashRecouvrement = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.accordionControlElement1 = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.mnuRecouvClient = new DevExpress.XtraBars.Navigation.AccordionControlElement();
            this.fluentDesignFormControl1 = new DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl();
            this.barSubItem1 = new DevExpress.XtraBars.BarSubItem();
            this.barButtonItem1 = new DevExpress.XtraBars.BarButtonItem();
            this.barButtonDisconnect = new DevExpress.XtraBars.BarButtonItem();
            this.barButtonTheme = new DevExpress.XtraBars.BarButtonItem();
            ((System.ComponentModel.ISupportInitialize)(this.popupMenu2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.accordionControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fluentDesignFormControl1)).BeginInit();
            this.SuspendLayout();
            // 
            // popupMenu2
            // 
            this.popupMenu2.Name = "popupMenu2";
            // 
            // ribbonPage2
            // 
            this.ribbonPage2.Name = "ribbonPage2";
            this.ribbonPage2.Text = "ribbonPage2";
            // 
            // splashScreenManager1
            // 
            this.splashScreenManager1.ClosingDelay = 500;
            // 
            // Container
            // 
            this.Container.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Container.Location = new System.Drawing.Point(297, 39);
            this.Container.Name = "Container";
            this.Container.Size = new System.Drawing.Size(912, 805);
            this.Container.TabIndex = 3;
            this.Container.Click += new System.EventHandler(this.Container_Click);
            // 
            // accordionControl1
            // 
            this.accordionControl1.Dock = System.Windows.Forms.DockStyle.Left;
            this.accordionControl1.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnuDSI,
            this.mnuDRH,
            this.mnuFINANCE,
            this.mnuGESCOM,
            this.mnuImportExport,
            this.mnuDashboard,
            this.accordionControlElement1});
            this.accordionControl1.Location = new System.Drawing.Point(0, 39);
            this.accordionControl1.Name = "accordionControl1";
            this.accordionControl1.ScrollBarMode = DevExpress.XtraBars.Navigation.ScrollBarMode.Touch;
            this.accordionControl1.Size = new System.Drawing.Size(297, 805);
            this.accordionControl1.TabIndex = 4;
            this.accordionControl1.ViewType = DevExpress.XtraBars.Navigation.AccordionControlViewType.HamburgerMenu;
            // 
            // mnuDSI
            // 
            this.mnuDSI.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnuVReleaseUser,
            this.mnuUpDatePrix});
            this.mnuDSI.Expanded = true;
            this.mnuDSI.Name = "mnuDSI";
            this.mnuDSI.Text = "DSI";
            // 
            // mnuVReleaseUser
            // 
            this.mnuVReleaseUser.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("mnuVReleaseUser.ImageOptions.SvgImage")));
            this.mnuVReleaseUser.Name = "mnuVReleaseUser";
            this.mnuVReleaseUser.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuVReleaseUser.Text = "Section admin";
            this.mnuVReleaseUser.Click += new System.EventHandler(this.accordionControlElement2_Click);
            // 
            // mnuUpDatePrix
            // 
            this.mnuUpDatePrix.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("mnuUpDatePrix.ImageOptions.SvgImage")));
            this.mnuUpDatePrix.Name = "mnuUpDatePrix";
            this.mnuUpDatePrix.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuUpDatePrix.Text = "Mise à jour des prix";
            this.mnuUpDatePrix.Click += new System.EventHandler(this.mnuUpDatePrix_Click_1);
            // 
            // mnuDRH
            // 
            this.mnuDRH.Name = "mnuDRH";
            this.mnuDRH.Text = "DRH";
            // 
            // mnuFINANCE
            // 
            this.mnuFINANCE.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnuMvgStk,
            this.accordionControlElement2});
            this.mnuFINANCE.Expanded = true;
            this.mnuFINANCE.Name = "mnuFINANCE";
            this.mnuFINANCE.Text = "FINANCE";
            // 
            // mnuMvgStk
            // 
            this.mnuMvgStk.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("mnuMvgStk.ImageOptions.SvgImage")));
            this.mnuMvgStk.Name = "mnuMvgStk";
            this.mnuMvgStk.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuMvgStk.Text = "Mouvements de stock";
            this.mnuMvgStk.Click += new System.EventHandler(this.accordionControlElement3_Click);
            // 
            // accordionControlElement2
            // 
            this.accordionControlElement2.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("accordionControlElement2.ImageOptions.SvgImage")));
            this.accordionControlElement2.Name = "accordionControlElement2";
            this.accordionControlElement2.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.accordionControlElement2.Text = "REGLEMENTS";
            this.accordionControlElement2.Visible = false;
            this.accordionControlElement2.Click += new System.EventHandler(this.accordionControlElement2_Click_1);
            // 
            // mnuGESCOM
            // 
            this.mnuGESCOM.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnuEditPrix});
            this.mnuGESCOM.Expanded = true;
            this.mnuGESCOM.Name = "mnuGESCOM";
            this.mnuGESCOM.Text = "GESTION COMMERCIALE";
            // 
            // mnuEditPrix
            // 
            this.mnuEditPrix.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("mnuEditPrix.ImageOptions.SvgImage")));
            this.mnuEditPrix.Name = "mnuEditPrix";
            this.mnuEditPrix.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuEditPrix.Text = "Edition nouveaux prix";
            this.mnuEditPrix.Click += new System.EventHandler(this.mnuEditPrix_Click);
            // 
            // mnuImportExport
            // 
            this.mnuImportExport.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.accordionControlElement3});
            this.mnuImportExport.Expanded = true;
            this.mnuImportExport.Name = "mnuImportExport";
            this.mnuImportExport.Text = "IMPORT";
            // 
            // accordionControlElement3
            // 
            this.accordionControlElement3.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("accordionControlElement3.ImageOptions.SvgImage")));
            this.accordionControlElement3.Name = "accordionControlElement3";
            this.accordionControlElement3.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.accordionControlElement3.Text = "Documents des achats";
            this.accordionControlElement3.Click += new System.EventHandler(this.accordionControlElement3_Click_1);
            // 
            // mnuDashboard
            // 
            this.mnuDashboard.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnuDashCIAL,
            this.mnuDashCPTA,
            this.mnuDashRecouvrement});
            this.mnuDashboard.Expanded = true;
            this.mnuDashboard.Name = "mnuDashboard";
            this.mnuDashboard.Text = "Dashboard";
            // 
            // mnuDashCIAL
            // 
            this.mnuDashCIAL.Name = "mnuDashCIAL";
            this.mnuDashCIAL.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuDashCIAL.Text = "Gestion commerciale";
            this.mnuDashCIAL.Click += new System.EventHandler(this.mnuDashCIAL_Click);
            // 
            // mnuDashCPTA
            // 
            this.mnuDashCPTA.Name = "mnuDashCPTA";
            this.mnuDashCPTA.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuDashCPTA.Text = "Comptabilité";
            this.mnuDashCPTA.Click += new System.EventHandler(this.mnuDashCPTA_Click);
            // 
            // mnuDashRecouvrement
            // 
            this.mnuDashRecouvrement.Name = "mnuDashRecouvrement";
            this.mnuDashRecouvrement.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuDashRecouvrement.Text = "Recouvrement";
            this.mnuDashRecouvrement.Click += new System.EventHandler(this.mnuDashRecouvrement_Click);
            // 
            // accordionControlElement1
            // 
            this.accordionControlElement1.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnuRecouvClient});
            this.accordionControlElement1.Expanded = true;
            this.accordionControlElement1.Name = "accordionControlElement1";
            this.accordionControlElement1.Text = "RECOUVREMENT";
            // 
            // mnuRecouvClient
            // 
            this.mnuRecouvClient.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("mnuRecouvClient.ImageOptions.SvgImage")));
            this.mnuRecouvClient.Name = "mnuRecouvClient";
            this.mnuRecouvClient.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
            this.mnuRecouvClient.Text = "Clients";
            this.mnuRecouvClient.Click += new System.EventHandler(this.mnuRecouvClient_Click);
            // 
            // fluentDesignFormControl1
            // 
            this.fluentDesignFormControl1.FluentDesignForm = this;
            this.fluentDesignFormControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.barSubItem1,
            this.barButtonItem1,
            this.barButtonTheme,
            this.barButtonDisconnect});
            this.fluentDesignFormControl1.Location = new System.Drawing.Point(0, 0);
            this.fluentDesignFormControl1.Name = "fluentDesignFormControl1";
            this.fluentDesignFormControl1.Size = new System.Drawing.Size(1209, 39);
            this.fluentDesignFormControl1.TabIndex = 5;
            this.fluentDesignFormControl1.TabStop = false;
            this.fluentDesignFormControl1.TitleItemLinks.Add(this.barButtonTheme);
            this.fluentDesignFormControl1.TitleItemLinks.Add(this.barSubItem1);
            this.fluentDesignFormControl1.Click += new System.EventHandler(this.fluentDesignFormControl1_Click);
            // 
            // barSubItem1
            // 
            this.barSubItem1.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right;
            this.barSubItem1.Caption = "user";
            this.barSubItem1.Id = 0;
            this.barSubItem1.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("barSubItem1.ImageOptions.SvgImage")));
            this.barSubItem1.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
            new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem1),
            new DevExpress.XtraBars.LinkPersistInfo(this.barButtonDisconnect)});
            this.barSubItem1.Name = "barSubItem1";
            this.barSubItem1.ShowMenuCaption = true;
            // 
            // barButtonItem1
            // 
            this.barButtonItem1.Caption = "Profil";
            this.barButtonItem1.Id = 1;
            this.barButtonItem1.Name = "barButtonItem1";
            this.barButtonItem1.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem1_ItemClick_1);
            // 
            // barButtonDisconnect
            // 
            this.barButtonDisconnect.Caption = "Déconnecter";
            this.barButtonDisconnect.Id = 0;
            this.barButtonDisconnect.Name = "barButtonDisconnect";
            this.barButtonDisconnect.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonDisconnect_ItemClick);
            // 
            // barButtonTheme
            // 
            this.barButtonTheme.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right;
            this.barButtonTheme.Id = 0;
            this.barButtonTheme.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("barButtonTheme.ImageOptions.SvgImage")));
            this.barButtonTheme.Name = "barButtonTheme";
            this.barButtonTheme.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonTheme_ItemClick);
            // 
            // FrmMdiParent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1209, 844);
            this.ControlContainer = this.Container;
            this.Controls.Add(this.Container);
            this.Controls.Add(this.accordionControl1);
            this.Controls.Add(this.fluentDesignFormControl1);
            this.FluentDesignFormControl = this.fluentDesignFormControl1;
            this.IconOptions.Image = ((System.Drawing.Image)(resources.GetObject("FrmMdiParent.IconOptions.Image")));
            this.Name = "FrmMdiParent";
            this.NavigationControl = this.accordionControl1;
            this.Text = "arBiochem";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FrmMdiParent_Load);
            ((System.ComponentModel.ISupportInitialize)(this.popupMenu2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.accordionControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fluentDesignFormControl1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage2;
        private DevExpress.XtraBars.PopupMenu popupMenu2;
        private DevExpress.XtraSplashScreen.SplashScreenManager splashScreenManager1;
        private new DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormContainer Container;
        private DevExpress.XtraBars.Navigation.AccordionControl accordionControl1;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuDSI;
        private DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl fluentDesignFormControl1;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuVReleaseUser;
        private DevExpress.XtraBars.BarSubItem barSubItem1;
        private DevExpress.XtraBars.BarButtonItem barButtonItem1;
        private DevExpress.XtraBars.BarButtonItem barButtonTheme;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuDRH;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuFINANCE;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuMvgStk;
        private DevExpress.XtraBars.BarButtonItem barButtonDisconnect;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuGESCOM;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuEditPrix;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuUpDatePrix;
        private DevExpress.XtraBars.Navigation.AccordionControlElement accordionControlElement2;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuImportExport;
        private DevExpress.XtraBars.Navigation.AccordionControlElement accordionControlElement3;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuDashboard;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuDashCIAL;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuDashCPTA;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuDashRecouvrement;
        private DevExpress.XtraBars.Navigation.AccordionControlElement accordionControlElement1;
        private DevExpress.XtraBars.Navigation.AccordionControlElement mnuRecouvClient;
    }
}