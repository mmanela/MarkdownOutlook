namespace MarkdownOutlook
{
    partial class AppointmentRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AppointmentRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.markdownGroup = this.Factory.CreateRibbonGroup();
            this.enableMarkdownMode = this.Factory.CreateRibbonToggleButton();
            this.renderMarkdown = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.markdownGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabAppointment";
            this.tab1.Groups.Add(this.markdownGroup);
            this.tab1.Label = "TabAppointment";
            this.tab1.Name = "tab1";
            // 
            // markdownGroup
            // 
            this.markdownGroup.Items.Add(this.enableMarkdownMode);
            this.markdownGroup.Items.Add(this.renderMarkdown);
            this.markdownGroup.Label = "Markdown";
            this.markdownGroup.Name = "markdownGroup";
            // 
            // enableMarkdownMode
            // 
            this.enableMarkdownMode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.enableMarkdownMode.Image = global::MarkdownOutlook.Properties.Resources.markdown;
            this.enableMarkdownMode.Label = "Enable Markdown Mode";
            this.enableMarkdownMode.Name = "enableMarkdownMode";
            this.enableMarkdownMode.ShowImage = true;
            this.enableMarkdownMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.enableMarkdownMode_Click);
            // 
            // renderMarkdown
            // 
            this.renderMarkdown.Image = global::MarkdownOutlook.Properties.Resources.markdown;
            this.renderMarkdown.Label = "Show Preview";
            this.renderMarkdown.Name = "renderMarkdown";
            this.renderMarkdown.ShowImage = true;
            this.renderMarkdown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RenderMarkdown_Click);
            // 
            // AppointmentRibbon
            // 
            this.Name = "AppointmentRibbon";
            this.RibbonType = "Microsoft.Outlook.Appointment";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AppointmentRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.markdownGroup.ResumeLayout(false);
            this.markdownGroup.PerformLayout();
        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup markdownGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton enableMarkdownMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton renderMarkdown;
    }

    partial class ThisRibbonCollection
    {
        internal AppointmentRibbon AppointmentRibbon
        {
            get { return this.GetRibbon<AppointmentRibbon>(); }
        }
    }
}
