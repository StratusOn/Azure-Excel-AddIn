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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AzureRibbon));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            this.AzureRibbonTab = this.Factory.CreateRibbonTab();
            this.AuthenticationRibbonGroup = this.Factory.CreateRibbonGroup();
            this.AuthTenantIdEditBox = this.Factory.CreateRibbonEditBox();
            this.TenantTypeDropDown = this.Factory.CreateRibbonDropDown();
            this.GetTokenButton = this.Factory.CreateRibbonButton();
            this.BillingAPIsRibbonGroup = this.Factory.CreateRibbonGroup();
            this.SubscriptionIdComboBox = this.Factory.CreateRibbonComboBox();
            this.TenantIdComboBox = this.Factory.CreateRibbonComboBox();
            this.ForceReAuthCheckBox = this.Factory.CreateRibbonCheckBox();
            this.StartDateEditBox = this.Factory.CreateRibbonEditBox();
            this.EndDateEditBox = this.Factory.CreateRibbonEditBox();
            this.AggregationGranularityDropDown = this.Factory.CreateRibbonDropDown();
            this.GetUsageReportButton = this.Factory.CreateRibbonButton();
            this.GetCspUsageReportButton = this.Factory.CreateRibbonButton();
            this.GetEaUsageReportButton = this.Factory.CreateRibbonButton();
            this.ApplicationIdComboBox = this.Factory.CreateRibbonComboBox();
            this.AppKeyComboBox = this.Factory.CreateRibbonComboBox();
            this.FillerLabel = this.Factory.CreateRibbonLabel();
            this.EnrollmentNumberComboBox = this.Factory.CreateRibbonComboBox();
            this.EaApiKeyComboBox = this.Factory.CreateRibbonComboBox();
            this.HelpGroup = this.Factory.CreateRibbonGroup();
            this.UpdateAddinButton = this.Factory.CreateRibbonButton();
            this.RateCardGroup = this.Factory.CreateRibbonGroup();
            this.GetRateCardButton = this.Factory.CreateRibbonButton();
            this.AzureRibbonTab.SuspendLayout();
            this.AuthenticationRibbonGroup.SuspendLayout();
            this.BillingAPIsRibbonGroup.SuspendLayout();
            this.HelpGroup.SuspendLayout();
            this.RateCardGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // AzureRibbonTab
            // 
            this.AzureRibbonTab.Groups.Add(this.AuthenticationRibbonGroup);
            this.AzureRibbonTab.Groups.Add(this.BillingAPIsRibbonGroup);
            this.AzureRibbonTab.Groups.Add(this.RateCardGroup);
            this.AzureRibbonTab.Groups.Add(this.HelpGroup);
            this.AzureRibbonTab.Label = "Azure";
            this.AzureRibbonTab.Name = "AzureRibbonTab";
            // 
            // AuthenticationRibbonGroup
            // 
            this.AuthenticationRibbonGroup.Items.Add(this.AuthTenantIdEditBox);
            this.AuthenticationRibbonGroup.Items.Add(this.TenantTypeDropDown);
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
            // TenantTypeDropDown
            // 
            ribbonDropDownItemImpl1.Label = "Standard";
            ribbonDropDownItemImpl1.ScreenTip = "Standard";
            ribbonDropDownItemImpl1.SuperTip = "Standard Azure account type.";
            ribbonDropDownItemImpl2.Label = "CSP";
            ribbonDropDownItemImpl2.ScreenTip = "CSP";
            ribbonDropDownItemImpl2.SuperTip = "Cloud Solution Provider (CSP) account type.";
            this.TenantTypeDropDown.Items.Add(ribbonDropDownItemImpl1);
            this.TenantTypeDropDown.Items.Add(ribbonDropDownItemImpl2);
            this.TenantTypeDropDown.Label = "Tenant Type";
            this.TenantTypeDropDown.Name = "TenantTypeDropDown";
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
            this.BillingAPIsRibbonGroup.Items.Add(this.ApplicationIdComboBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.AppKeyComboBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.FillerLabel);
            this.BillingAPIsRibbonGroup.Items.Add(this.EnrollmentNumberComboBox);
            this.BillingAPIsRibbonGroup.Items.Add(this.EaApiKeyComboBox);
            this.BillingAPIsRibbonGroup.Label = "Azure Usage APIs";
            this.BillingAPIsRibbonGroup.Name = "BillingAPIsRibbonGroup";
            // 
            // SubscriptionIdComboBox
            // 
            this.SubscriptionIdComboBox.Items.Add(ribbonDropDownItemImpl3);
            this.SubscriptionIdComboBox.Label = "Subscription Id";
            this.SubscriptionIdComboBox.Name = "SubscriptionIdComboBox";
            this.SubscriptionIdComboBox.ScreenTip = "Subscription Id";
            this.SubscriptionIdComboBox.SuperTip = "Subscription id for which to get aggregate usage, in the form xxxxxxxx-xxxx-xxxx-" +
    "xxxx-xxxxxxxxxxxx.";
            this.SubscriptionIdComboBox.Text = null;
            // 
            // TenantIdComboBox
            // 
            this.TenantIdComboBox.Items.Add(ribbonDropDownItemImpl4);
            this.TenantIdComboBox.Label = "Tenant Id";
            this.TenantIdComboBox.Name = "TenantIdComboBox";
            this.TenantIdComboBox.ScreenTip = "Tenant Id";
            this.TenantIdComboBox.SuperTip = "The user\'s tenant id (standard and CSP) or the customer tenant id (EA), in the fo" +
    "rm xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.";
            this.TenantIdComboBox.Text = null;
            // 
            // ForceReAuthCheckBox
            // 
            this.ForceReAuthCheckBox.Checked = true;
            this.ForceReAuthCheckBox.Label = "Show Authentication Dialog";
            this.ForceReAuthCheckBox.Name = "ForceReAuthCheckBox";
            this.ForceReAuthCheckBox.ScreenTip = "Show Authentication Dialog";
            this.ForceReAuthCheckBox.SuperTip = "Uncheck to use cached credentials. Keep checked to always show the user authentic" +
    "ation page and get fresh credentials.";
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
            // AggregationGranularityDropDown
            // 
            ribbonDropDownItemImpl5.Label = "Daily";
            ribbonDropDownItemImpl5.Tag = "Daily";
            ribbonDropDownItemImpl6.Label = "Hourly";
            ribbonDropDownItemImpl6.Tag = "Hourly";
            this.AggregationGranularityDropDown.Items.Add(ribbonDropDownItemImpl5);
            this.AggregationGranularityDropDown.Items.Add(ribbonDropDownItemImpl6);
            this.AggregationGranularityDropDown.Label = "Aggregation Granularity";
            this.AggregationGranularityDropDown.Name = "AggregationGranularityDropDown";
            this.AggregationGranularityDropDown.ScreenTip = "Aggregation Granularity";
            this.AggregationGranularityDropDown.SuperTip = "Data granularity (Daily or Hourly). The default is Daily.";
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
            // ApplicationIdComboBox
            // 
            this.ApplicationIdComboBox.Items.Add(ribbonDropDownItemImpl7);
            this.ApplicationIdComboBox.Label = "App Id";
            this.ApplicationIdComboBox.Name = "ApplicationIdComboBox";
            this.ApplicationIdComboBox.ScreenTip = "Application Id";
            this.ApplicationIdComboBox.SuperTip = "An optional application id (GUID) of an AAD client application.";
            this.ApplicationIdComboBox.Text = null;
            // 
            // AppKeyComboBox
            // 
            this.AppKeyComboBox.Items.Add(ribbonDropDownItemImpl8);
            this.AppKeyComboBox.Label = "App Key";
            this.AppKeyComboBox.Name = "AppKeyComboBox";
            this.AppKeyComboBox.ScreenTip = "App Key";
            this.AppKeyComboBox.SuperTip = "Application key (secret) for the AAD application.";
            this.AppKeyComboBox.Text = null;
            // 
            // FillerLabel
            // 
            this.FillerLabel.Label = " ";
            this.FillerLabel.Name = "FillerLabel";
            // 
            // EnrollmentNumberComboBox
            // 
            this.EnrollmentNumberComboBox.Items.Add(ribbonDropDownItemImpl9);
            this.EnrollmentNumberComboBox.Label = "EA Enrollment Number";
            this.EnrollmentNumberComboBox.Name = "EnrollmentNumberComboBox";
            this.EnrollmentNumberComboBox.ScreenTip = "EA Enrollment Number";
            this.EnrollmentNumberComboBox.SuperTip = "The Enrollment Number for the EA for which usage data is to be collected.";
            this.EnrollmentNumberComboBox.Text = null;
            // 
            // EaApiKeyComboBox
            // 
            this.EaApiKeyComboBox.Items.Add(ribbonDropDownItemImpl10);
            this.EaApiKeyComboBox.Label = "EA API Key";
            this.EaApiKeyComboBox.Name = "EaApiKeyComboBox";
            this.EaApiKeyComboBox.ScreenTip = "EA API Key";
            this.EaApiKeyComboBox.SuperTip = "An EA API Key (generated in the EA portal, http://ea.azure.com) is required for g" +
    "etting an EA Usage Report.";
            this.EaApiKeyComboBox.Text = null;
            // 
            // HelpGroup
            // 
            this.HelpGroup.Items.Add(this.UpdateAddinButton);
            this.HelpGroup.Label = "Help";
            this.HelpGroup.Name = "HelpGroup";
            // 
            // UpdateAddinButton
            // 
            this.UpdateAddinButton.Label = "Update Add-in";
            this.UpdateAddinButton.Name = "UpdateAddinButton";
            this.UpdateAddinButton.ScreenTip = "Update Add-in";
            this.UpdateAddinButton.SuperTip = "Run the ClickOnce installer to update the add-in if a new version is available.";
            this.UpdateAddinButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateAddinButton_Click);
            // 
            // RateCardGroup
            // 
            this.RateCardGroup.Items.Add(this.GetRateCardButton);
            this.RateCardGroup.Label = "RateCard";
            this.RateCardGroup.Name = "RateCardGroup";
            // 
            // GetRateCardButton
            // 
            this.GetRateCardButton.Image = global::ExcelAddIn1.Properties.Resources.BillingHub;
            this.GetRateCardButton.Label = "Get RateCard";
            this.GetRateCardButton.Name = "GetRateCardButton";
            this.GetRateCardButton.ShowImage = true;
            this.GetRateCardButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetRateCardButton_Click);
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
            this.HelpGroup.ResumeLayout(false);
            this.HelpGroup.PerformLayout();
            this.RateCardGroup.ResumeLayout(false);
            this.RateCardGroup.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown TenantTypeDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup HelpGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateAddinButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox ApplicationIdComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox AppKeyComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel FillerLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RateCardGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetRateCardButton;
    }

    partial class ThisRibbonCollection
    {
        internal AzureRibbon Ribbon1
        {
            get { return this.GetRibbon<AzureRibbon>(); }
        }
    }
}
