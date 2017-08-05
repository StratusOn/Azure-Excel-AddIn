namespace ExcelAddIn1
{
    partial class AzureRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AzureRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AzureRibbon));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.AzureRibbonTab = this.Factory.CreateRibbonTab();
            this.AuthenticationRibbonGroup = this.Factory.CreateRibbonGroup();
            this.AuthTenantIdEditBox = this.Factory.CreateRibbonEditBox();
            this.GetTokenButton = this.Factory.CreateRibbonButton();
            this.BillingAPIsRibbonGroup = this.Factory.CreateRibbonGroup();
            this.SubscriptionIdEditBox = this.Factory.CreateRibbonEditBox();
            this.TenantIdEditBox = this.Factory.CreateRibbonEditBox();
            this.AggregationGranularityDropDown = this.Factory.CreateRibbonDropDown();
            this.StartDateEditBox = this.Factory.CreateRibbonEditBox();
            this.EndDateEditBox = this.Factory.CreateRibbonEditBox();
            this.ForceReAuthCheckBox = this.Factory.CreateRibbonCheckBox();
            this.GetUsageReportButton = this.Factory.CreateRibbonButton();
            this.AzureRibbonTab.SuspendLayout();
            this.AuthenticationRibbonGroup.SuspendLayout();
            this.BillingAPIsRibbonGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // AzureRibbonTab
            // 
            this.AzureRibbonTab.Groups.Add(this.AuthenticationRibbonGroup);
            this.AzureRibbonTab.Groups.Add(this.BillingAPIsRibbonGroup);
            this.AzureRibbonTab.Label = "Azure";
            this.AzureRibbonTab.Name = "AzureRibbonTab";
            // 
            // AuthenticationRibbonGroup
            // 
            this.AuthenticationRibbonGroup.Items.Add(this.AuthTenantIdEditBox);
            this.AuthenticationRibbonGroup.Items.Add(this.GetTokenButton);
            this.AuthenticationRibbonGroup.Label = "Authentication";
            this.AuthenticationRibbonGroup.Name = "AuthenticationRibbonGroup";
            // 
            // AuthTenantIdEditBox
            // 
            this.AuthTenantIdEditBox.Label = "Tenant Id";
            this.AuthTenantIdEditBox.Name = "AuthTenantIdEditBox";
            this.AuthTenantIdEditBox.Text = null;
            // 
            // GetTokenButton
            // 
            this.GetTokenButton.Image = ((System.Drawing.Image)(resources.GetObject("GetTokenButton.Image")));
            this.GetTokenButton.Label = "Get Authentication Token";
            this.GetTokenButton.Name = "GetTokenButton";
            this.GetTokenButton.ShowImage = true;
            this.GetTokenButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetTokenButton_Click);
            // 
            // BillingAPIsRibbonGroup
            // 
            this.BillingAPIsRibbonGroup.Items.Add(this.SubscriptionIdEditBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.TenantIdEditBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.AggregationGranularityDropDown);
            this.BillingAPIsRibbonGroup.Items.Add(this.StartDateEditBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.EndDateEditBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.ForceReAuthCheckBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.GetUsageReportButton);
            this.BillingAPIsRibbonGroup.Label = "Consumption APIs";
            this.BillingAPIsRibbonGroup.Name = "BillingAPIsRibbonGroup";
            // 
            // SubscriptionIdEditBox
            // 
            this.SubscriptionIdEditBox.Label = "Subscription Id";
            this.SubscriptionIdEditBox.Name = "SubscriptionIdEditBox";
            this.SubscriptionIdEditBox.SuperTip = "Subscription id for which to get aggregate usage, in the form xxxxxxxx-xxxx-xxxx-" +
    "xxxx-xxxxxxxxxxxx.";
            this.SubscriptionIdEditBox.Text = null;
            // 
            // TenantIdEditBox
            // 
            this.TenantIdEditBox.Label = "Tenant Id";
            this.TenantIdEditBox.Name = "TenantIdEditBox";
            this.TenantIdEditBox.Text = null;
            // 
            // AggregationGranularityDropDown
            // 
            ribbonDropDownItemImpl1.Label = "Daily";
            ribbonDropDownItemImpl1.Tag = "Daily";
            ribbonDropDownItemImpl2.Label = "Hourly";
            ribbonDropDownItemImpl2.Tag = "Hourly";
            this.AggregationGranularityDropDown.Items.Add(ribbonDropDownItemImpl1);
            this.AggregationGranularityDropDown.Items.Add(ribbonDropDownItemImpl2);
            this.AggregationGranularityDropDown.Label = "Aggregation Granularity";
            this.AggregationGranularityDropDown.Name = "AggregationGranularityDropDown";
            // 
            // StartDateEditBox
            // 
            this.StartDateEditBox.Label = "Report Start Date";
            this.StartDateEditBox.Name = "StartDateEditBox";
            this.StartDateEditBox.SuperTip = "Report Start Date (yyyy-mm-dd)";
            this.StartDateEditBox.Text = null;
            // 
            // EndDateEditBox
            // 
            this.EndDateEditBox.Label = "Report End Date";
            this.EndDateEditBox.Name = "EndDateEditBox";
            this.EndDateEditBox.SuperTip = "Report End Date (yyyy-mm-dd)";
            this.EndDateEditBox.Text = null;
            // 
            // ForceReAuthCheckBox
            // 
            this.ForceReAuthCheckBox.Checked = true;
            this.ForceReAuthCheckBox.Label = "Force Re-Authentication";
            this.ForceReAuthCheckBox.Name = "ForceReAuthCheckBox";
            // 
            // GetUsageReportButton
            // 
            this.GetUsageReportButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GetUsageReportButton.Image = ((System.Drawing.Image)(resources.GetObject("GetUsageReportButton.Image")));
            this.GetUsageReportButton.Label = "Get Usage Report";
            this.GetUsageReportButton.Name = "GetUsageReportButton";
            this.GetUsageReportButton.ShowImage = true;
            this.GetUsageReportButton.SuperTip = resources.GetString("GetUsageReportButton.SuperTip");
            this.GetUsageReportButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetUsageReportButton_Click);
            // 
            // AzureRibbon
            // 
            this.Name = "AzureRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.AzureRibbonTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AzureRibbonTab_Load);
            this.AzureRibbonTab.ResumeLayout(false);
            this.AzureRibbonTab.PerformLayout();
            this.AuthenticationRibbonGroup.ResumeLayout(false);
            this.AuthenticationRibbonGroup.PerformLayout();
            this.BillingAPIsRibbonGroup.ResumeLayout(false);
            this.BillingAPIsRibbonGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private Microsoft.Office.Tools.Ribbon.RibbonTab AzureRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup BillingAPIsRibbonGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetUsageReportButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SubscriptionIdEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox StartDateEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EndDateEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown AggregationGranularityDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ForceReAuthCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox TenantIdEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AuthenticationRibbonGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox AuthTenantIdEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetTokenButton;
    }

    partial class ThisRibbonCollection
    {
        internal AzureRibbon Ribbon1
        {
            get { return this.GetRibbon<AzureRibbon>(); }
        }
    }
}
