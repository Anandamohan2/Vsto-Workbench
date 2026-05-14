namespace PowerPointVstoAdd_In
{
    partial class PresentationInfoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PresentationInfoRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary>
        /// Clean up resources.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing &&
                (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.PresentationInfo = this.Factory.CreateRibbonGroup();
            this.btnReadInfo = this.Factory.CreateRibbonButton();

            this.tab1.SuspendLayout();
            this.PresentationInfo.SuspendLayout();
            this.SuspendLayout();

            // tab1
            this.tab1.ControlId.ControlIdType =
                Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;

            this.tab1.Groups.Add(this.PresentationInfo);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";

            // PresentationInfo
            this.PresentationInfo.Items.Add(this.btnReadInfo);
            this.PresentationInfo.Label = "Presentation Info";
            this.PresentationInfo.Name = "PresentationInfo";

            // btnReadInfo
            this.btnReadInfo.Label =
                "Read Presentation Info";

            this.btnReadInfo.Name =
                "btnReadInfo";

            this.btnReadInfo.Click +=
                new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                    this.btnReadInfo_Click);

            // PresentationInfoRibbon
            this.Name = "PresentationInfoRibbon";
            this.RibbonType =
                "Microsoft.PowerPoint.Presentation";

            this.Tabs.Add(this.tab1);

            this.Load +=
                new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(
                    this.PresentationInfoRibbon_Load);

            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();

            this.PresentationInfo.ResumeLayout(false);
            this.PresentationInfo.PerformLayout();

            this.ResumeLayout(false);
        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup PresentationInfo;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadInfo;
    }

    partial class ThisRibbonCollection
    {
        internal PresentationInfoRibbon PresentationInfoRibbon
        {
            get
            {
                return this.GetRibbon<PresentationInfoRibbon>();
            }
        }
    }
}