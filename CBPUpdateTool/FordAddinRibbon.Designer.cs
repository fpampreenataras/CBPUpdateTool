namespace CBPUpdateTool
{
    partial class FordAddinRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FordAddinRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.UpdateFromAras = this.Factory.CreateRibbonButton();
            this.CompleteTaskButton = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.WorkflowLbl = this.Factory.CreateRibbonLabel();
            this.WorkflowNameRibbonLabel = this.Factory.CreateRibbonLabel();
            this.box2 = this.Factory.CreateRibbonBox();
            this.ActivityLbl = this.Factory.CreateRibbonLabel();
            this.ActivityRibbonLabel = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Ford Addin";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.UpdateFromAras);
            this.group1.Items.Add(this.CompleteTaskButton);
            this.group1.Items.Add(this.separator1);
            this.group1.Label = "Ford Add-in";
            this.group1.Name = "group1";
            // 
            // UpdateFromAras
            // 
            this.UpdateFromAras.Label = "Update CBP Arch";
            this.UpdateFromAras.Name = "UpdateFromAras";
            this.UpdateFromAras.OfficeImageId = "TableSharePointListsRefreshList";
            this.UpdateFromAras.ShowImage = true;
            this.UpdateFromAras.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateFromAras_Click);
            // 
            // CompleteTaskButton
            // 
            this.CompleteTaskButton.Label = "Complete Activity";
            this.CompleteTaskButton.Name = "CompleteTaskButton";
            this.CompleteTaskButton.OfficeImageId = "AcceptInvitation";
            this.CompleteTaskButton.ShowImage = true;
            this.CompleteTaskButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Complete_Activity_click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.box1);
            this.group2.Items.Add(this.box2);
            this.group2.Label = "Workflow Status";
            this.group2.Name = "group2";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.WorkflowLbl);
            this.box1.Items.Add(this.WorkflowNameRibbonLabel);
            this.box1.Name = "box1";
            // 
            // WorkflowLbl
            // 
            this.WorkflowLbl.Label = "Workflow";
            this.WorkflowLbl.Name = "WorkflowLbl";
            // 
            // WorkflowNameRibbonLabel
            // 
            this.WorkflowNameRibbonLabel.Label = " ";
            this.WorkflowNameRibbonLabel.Name = "WorkflowNameRibbonLabel";
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.ActivityLbl);
            this.box2.Items.Add(this.ActivityRibbonLabel);
            this.box2.Name = "box2";
            // 
            // ActivityLbl
            // 
            this.ActivityLbl.Label = "Activity";
            this.ActivityLbl.Name = "ActivityLbl";
            // 
            // ActivityRibbonLabel
            // 
            this.ActivityRibbonLabel.Label = " ";
            this.ActivityRibbonLabel.Name = "ActivityRibbonLabel";
            // 
            // FordAddinRibbon
            // 
            this.Name = "FordAddinRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateFromAras;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CompleteTaskButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel WorkflowLbl;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel ActivityLbl;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel WorkflowNameRibbonLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel ActivityRibbonLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
    }

    partial class ThisRibbonCollection
    {
        internal FordAddinRibbon Ribbon1
        {
            get { return this.GetRibbon<FordAddinRibbon>(); }
        }
    }
}
