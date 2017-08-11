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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AzureRibbon));
            this.AzureRibbonTab = this.Factory.CreateRibbonTab();
            this.AuthenticationRibbonGroup = this.Factory.CreateRibbonGroup();
            this.AuthTenantIdEditBox = this.Factory.CreateRibbonEditBox();
            this.GetTokenButton = this.Factory.CreateRibbonButton();
            this.BillingAPIsRibbonGroup = this.Factory.CreateRibbonGroup();
            this.SubscriptionIdComboBox = this.Factory.CreateRibbonComboBox();
            this.TenantIdComboBox = this.Factory.CreateRibbonComboBox();
            this.AggregationGranularityDropDown = this.Factory.CreateRibbonDropDown();
            this.StartDateEditBox = this.Factory.CreateRibbonEditBox();
            this.EndDateEditBox = this.Factory.CreateRibbonEditBox();
            this.GetUsageReportButton = this.Factory.CreateRibbonButton();
            this.GetCspUsageReportButton = this.Factory.CreateRibbonButton();
            this.GetEaUsageReportButton = this.Factory.CreateRibbonButton();
            this.ForceReAuthCheckBox = this.Factory.CreateRibbonCheckBox();
            this.EaApiKeyComboBox = this.Factory.CreateRibbonComboBox();
            this.EnrollmentNumberComboBox = this.Factory.CreateRibbonComboBox();
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
            this.AuthTenantIdEditBox.ScreenTip = "Tenant Id";
            this.AuthTenantIdEditBox.SuperTip = "The user\'s tenant id (standard and CSP) or the customer tenant id (EA), in the fo" +
    "rm xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.";
            this.AuthTenantIdEditBox.Text = null;
            // 
            // GetTokenButton
            // 
            this.GetTokenButton.Image = global::ExcelAddIn1.Properties.Resources.Azure_Acitve_Directory_Access_Control;
            this.GetTokenButton.Label = "Get Authentication Token";
            this.GetTokenButton.Name = "GetTokenButton";
            this.GetTokenButton.ShowImage = true;
            this.GetTokenButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetTokenButton_Click);
            // 
            // BillingAPIsRibbonGroup
            // 
            this.BillingAPIsRibbonGroup.Items.Add(this.SubscriptionIdComboBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.TenantIdComboBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.ForceReAuthCheckBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.StartDateEditBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.EndDateEditBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.AggregationGranularityDropDown);
            this.BillingAPIsRibbonGroup.Items.Add(this.GetUsageReportButton);
            this.BillingAPIsRibbonGroup.Items.Add(this.GetCspUsageReportButton);
            this.BillingAPIsRibbonGroup.Items.Add(this.GetEaUsageReportButton);
            this.BillingAPIsRibbonGroup.Items.Add(this.EnrollmentNumberComboBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.EaApiKeyComboBox);
            this.BillingAPIsRibbonGroup.Label = "Azure Usage APIs";
            this.BillingAPIsRibbonGroup.Name = "BillingAPIsRibbonGroup";
            // 
            // SubscriptionIdComboBox
            // 
            this.SubscriptionIdComboBox.Label = "Subscription Id";
            this.SubscriptionIdComboBox.Name = "SubscriptionIdComboBox";
            this.SubscriptionIdComboBox.ScreenTip = "Subscription Id";
            this.SubscriptionIdComboBox.SuperTip = "Subscription id for which to get aggregate usage, in the form xxxxxxxx-xxxx-xxxx-" +
    "xxxx-xxxxxxxxxxxx.";
            this.SubscriptionIdComboBox.Text = null;
            // 
            // TenantIdComboBox
            // 
            this.TenantIdComboBox.Label = "Tenant Id";
            this.TenantIdComboBox.Name = "TenantIdComboBox";
            this.TenantIdComboBox.ScreenTip = "Tenant Id";
            this.TenantIdComboBox.SuperTip = "The user\'s tenant id (standard and CSP) or the customer tenant id (EA), in the fo" +
    "rm xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.";
            this.TenantIdComboBox.Text = null;
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
            this.AggregationGranularityDropDown.ScreenTip = "Aggregation Granularity";
            this.AggregationGranularityDropDown.SuperTip = "Data granularity (Daily or Hourly). The default is Daily.";
            // 
            // StartDateEditBox
            // 
            this.StartDateEditBox.Label = "Report Start Date";
            this.StartDateEditBox.Name = "StartDateEditBox";
            this.StartDateEditBox.ScreenTip = "Report Start Date";
            this.StartDateEditBox.SuperTip = "Report Start Date (yyyy-mm-dd). It can include a time portion for standard and CS" +
    "P accounts.";
            this.StartDateEditBox.Text = null;
            // 
            // EndDateEditBox
            // 
            this.EndDateEditBox.Label = "Report End Date";
            this.EndDateEditBox.Name = "EndDateEditBox";
            this.EndDateEditBox.ScreenTip = "Report End Date";
            this.EndDateEditBox.SuperTip = "Report End Date (yyyy-mm-dd).  It can include a time portion for standard and CSP" +
    " accounts.";
            this.EndDateEditBox.Text = null;
            // 
            // GetUsageReportButton
            // 
            this.GetUsageReportButton.Image = global::ExcelAddIn1.Properties.Resources.BillingHub;
            this.GetUsageReportButton.Label = "Get Usage Report (Standard)";
            this.GetUsageReportButton.Name = "GetUsageReportButton";
            this.GetUsageReportButton.ScreenTip = "Get Usage Report (Standard)";
            this.GetUsageReportButton.ShowImage = true;
            this.GetUsageReportButton.SuperTip = resources.GetString("GetUsageReportButton.SuperTip");
            this.GetUsageReportButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetUsageReportButton_Click);
            // 
            // GetCspUsageReportButton
            // 
            this.GetCspUsageReportButton.Image = global::ExcelAddIn1.Properties.Resources.BillingHub;
            this.GetCspUsageReportButton.Label = "Get Usage Report (CSP)";
            this.GetCspUsageReportButton.Name = "GetCspUsageReportButton";
            this.GetCspUsageReportButton.ScreenTip = "Get Usage Report (CSP)";
            this.GetCspUsageReportButton.ShowImage = true;
            this.GetCspUsageReportButton.SuperTip = resources.GetString("GetCspUsageReportButton.SuperTip");
            this.GetCspUsageReportButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetCspUsageReportButton_Click);
            // 
            // GetEaUsageReportButton
            // 
            this.GetEaUsageReportButton.Image = global::ExcelAddIn1.Properties.Resources.BillingHub;
            this.GetEaUsageReportButton.Label = "Get Usage Report (EA)";
            this.GetEaUsageReportButton.Name = "GetEaUsageReportButton";
            this.GetEaUsageReportButton.ScreenTip = "Get Usage Report (EA)";
            this.GetEaUsageReportButton.ShowImage = true;
            this.GetEaUsageReportButton.SuperTip = "Enter a tenant id, an enrollment id, a report start date (yyyy-mm-dd), and a repo" +
    "rt end date before clicking on this button. Check the Force Re-authnetication to" +
    " always get a fresh token.";
            this.GetEaUsageReportButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetEaUsageReportButton_Click);
            // 
            // ForceReAuthCheckBox
            // 
            this.ForceReAuthCheckBox.Checked = true;
            this.ForceReAuthCheckBox.Label = "Force Re-Authentication";
            this.ForceReAuthCheckBox.Name = "ForceReAuthCheckBox";
            this.ForceReAuthCheckBox.ScreenTip = "Force Re-Authentication";
            this.ForceReAuthCheckBox.SuperTip = "Uncheck to use cached credentials. Keep checked to always gets fresh credentials." +
    "";
            // 
            // EaApiKeyComboBox
            // 
            this.EaApiKeyComboBox.Label = "EA API Key";
            this.EaApiKeyComboBox.Name = "EaApiKeyComboBox";
            this.EaApiKeyComboBox.ScreenTip = "EA API Key";
            this.EaApiKeyComboBox.SuperTip = "An EA API Key (generated in the EA portal, http://ea.azure.com) is required for g" +
    "etting an EA Usage Report.";
            // 
            // EnrollmentNumberComboBox
            // 
            this.EnrollmentNumberComboBox.Label = "EA Enrollment Number";
            this.EnrollmentNumberComboBox.Name = "EnrollmentNumberComboBox";
            this.EnrollmentNumberComboBox.ScreenTip = "EA Enrollment Number";
            this.EnrollmentNumberComboBox.SuperTip = "The Enrollment Number for the EA for which usage data is to be collected.";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox StartDateEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EndDateEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown AggregationGranularityDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ForceReAuthCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AuthenticationRibbonGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox AuthTenantIdEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetTokenButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox SubscriptionIdComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox TenantIdComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetCspUsageReportButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetEaUsageReportButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox EaApiKeyComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox EnrollmentNumberComboBox;
    }

    partial class ThisRibbonCollection
    {
        internal AzureRibbon Ribbon1
        {
            get { return this.GetRibbon<AzureRibbon>(); }
        }
    }
}
