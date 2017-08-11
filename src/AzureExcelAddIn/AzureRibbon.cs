using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class AzureRibbon
    {
        private const int DefaultChunkSize = 1000; // CSP allows specifying chunk size. Default is 1000. Max is 1000.
        private const int MaxContinuationLinks = 500;
        private const string AddinInstallUrl = "http://billingtools.azurewebsites.net/excel/install/setup.exe";

        private readonly string[] HeaderCaptions = {
            "Usage Start Time (UTC)", "Usage End Time (UTC)", "Id", "Name", "Type", "subscription Id", "Meter Id", "Meter Name",
            "Meter Category", "Meter Sub-Category", "Quantity", "Unit", "Tags", "Info Fields (legacy format)", "Instance Data (new format)"
        };
        private readonly string[] HeaderCaptionsCsp = {
            "Usage Start Time (UTC)", "Usage End Time (UTC)", "Meter Id", "Meter Name", "Meter Category", "Meter Sub-Category", "Meter Region",
            "Quantity", "Unit", "Tags", "Info Fields (legacy format)", "Instance Data (new format), Attributes"
        };
        private readonly string[] HeaderCaptionsEa = {
            "Usage Time (UTC)", "Account Id", "Account Name", "Product Id", "Product", "Resource Location Id", "Resource Location", "Consumed Service Id", "Consumed Service", "Department Id", "Department Name",
            "Account Owner Email", "Service Administrator Id", "Subscription Id", "Subscription Guid", "Subscription Name", "Tags",
            "Meter Id", "Meter Name", "Meter Category", "Meter Sub-Category", "Meter Region", "Consumed Quantity", "Unit of Measure", "Resource Rate", "Cost",
            "Instance Id", "Service Info 1", "Service Info 2", "Additional Info", "Store Service Identifier", "Cost Center", "Resource Group"
        };

        private void AzureRibbonTab_Load(object sender, RibbonUIEventArgs e)
        {
            var today = DateTime.Today;
            var yesterday = today.AddDays(-1);
            this.StartDateEditBox.Text = $"{yesterday.Year}-{yesterday.Month:0#}-{yesterday.Day:0#}";
            this.EndDateEditBox.Text = $"{today.Year}-{today.Month:0#}-{today.Day:0#}";

            this.HydrateFromPersistedData();
        }

        private void GetTokenButton_Click(object sender, RibbonControlEventArgs e)
        {
            var tenantId = this.AuthTenantIdEditBox.Text;
            if (string.IsNullOrWhiteSpace(tenantId))
            {
                MessageBox.Show($"ERROR: Tenant Id must be specified.", "Get Authentication Token", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                UsageApi usageApi = this.TenantTypeDropDown.SelectedItemIndex == 1
                    ? UsageApi.CloudSolutionProvider
                    : UsageApi.Standard;
                string token = AuthUtils.GetAuthorizationHeader(tenantId, true, usageApi);

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show($"ERROR: Failed to acquire a token. Verify you entered the right credentials and the correct Tenant Id and try again.", "Get Authentication Token", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                MessageBox.Show($"{token}", "Your Authentication Token (Press CTRL+C to copy)", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to acquire token: {ex.Message}\r\n\r\n{ex.StackTrace}\r\n", "Get Authentication Token", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void GetUsageReportButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!this.ValidateUsageReportInput(UsageApi.Standard))
            {
                return;
            }

            var tenantId = this.TenantIdComboBox.Text;

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Authenticating...";

                string token = AuthUtils.GetAuthorizationHeader(tenantId, this.ForceReAuthCheckBox.Checked, UsageApi.Standard);

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show(
                        $"ERROR: Failed to acquire a token. Verify you entered the right credentials, the correct Tenant Id and Subscription Id, and make sure 'Force Re-Authentication' is checked and try again.",
                        "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Globals.ThisAddIn.Application.StatusBar = "Getting usage report (standard)...";

                var subscriptionId = this.SubscriptionIdComboBox.Text;
                var reportStartDate = this.StartDateEditBox.Text;
                var reportEndDate = this.EndDateEditBox.Text;
                var aggregationGranularity = (string)this.AggregationGranularityDropDown.SelectedItem.Tag;
                var showDetails = "true"; // this.ShowDetailsCheckBox.Checked.ToString();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int headerRowNumber = 9;
                int rowNumber = headerRowNumber + 2;
                Excel.Worksheet currentActiveWorksheet = null;
                int currentContinuationCount = 0;
                UsageAggregates usageAggregates = await BillingUtils.GetUsageAggregatesStandard(token, subscriptionId, reportStartDate, reportEndDate, aggregationGranularity, showDetails);

                do
                {
                    if (usageAggregates == null)
                    {
                        MessageBox.Show(
                            $"ERROR: Failed to get usage report. Verify the correct parameters were provided for Subscription Id, Start Date, and End Date and try again.",
                            "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (currentContinuationCount == 0)
                    {
                        // Add a fresh worksheet.
                        Excel.Worksheet previousActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
                        currentActiveWorksheet = Globals.ThisAddIn.Application.Worksheets.Add(previousActiveWorksheet);
                        this.PrintUsageAggregatesHeader(startColumnNumber, headerRowNumber, currentActiveWorksheet, UsageApi.Standard);
                    }

                    this.PrintUsageAggregatesReport(startColumnNumber, rowNumber, usageAggregates, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.value.Length;

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.nextLink;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    string content = await BillingUtils.GetUsageAggregates(token, continuationLink);
                    usageAggregates = JsonConvert.DeserializeObject<UsageAggregates>(content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                this.FormatTags(UsageApi.Standard, rowNumber, currentActiveWorksheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to get usage report: {ex.Message}\r\n\r\n\r\n{ex.StackTrace}\r\n",
                    "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private async void GetCspUsageReportButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!this.ValidateUsageReportInput(UsageApi.CloudSolutionProvider))
            {
                return;
            }

            var tenantId = this.TenantIdComboBox.Text;

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Authenticating...";

                string token = AuthUtils.GetAuthorizationHeader(tenantId, this.ForceReAuthCheckBox.Checked, UsageApi.CloudSolutionProvider);

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show(
                        $"ERROR: Failed to acquire a token. Verify you entered the right credentials, the correct Tenant Id and Subscription Id, and make sure 'Force Re-Authentication' is checked and try again.",
                        "Get Usage Report (CSP)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Globals.ThisAddIn.Application.StatusBar = "Getting usage report (CSP)...";

                var subscriptionId = this.SubscriptionIdComboBox.Text;
                var reportStartDate = this.StartDateEditBox.Text;
                var reportEndDate = this.EndDateEditBox.Text;
                var aggregationGranularity = (string)this.AggregationGranularityDropDown.SelectedItem.Tag;
                var showDetails = "true"; // this.ShowDetailsCheckBox.Checked.ToString();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int headerRowNumber = 9;
                int rowNumber = headerRowNumber + 2;
                Excel.Worksheet currentActiveWorksheet = null;
                int currentContinuationCount = 0;
                int chunkSize = DefaultChunkSize;
                CspUsageAggregates usageAggregates = await BillingUtils.GetUsageAggregatesCsp(token, subscriptionId, tenantId, reportStartDate, reportEndDate, aggregationGranularity, showDetails, chunkSize);

                do
                {
                    if (usageAggregates == null)
                    {
                        MessageBox.Show(
                            $"ERROR: Failed to get usage report. Verify the correct parameters were provided for Subscription Id, Start Date, and End Date and try again.",
                            "Get Usage Report (CSP)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (currentContinuationCount == 0)
                    {
                        // Add a fresh worksheet.
                        Excel.Worksheet previousActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
                        currentActiveWorksheet = Globals.ThisAddIn.Application.Worksheets.Add(previousActiveWorksheet);
                        this.PrintUsageAggregatesHeader(startColumnNumber, headerRowNumber, currentActiveWorksheet, UsageApi.CloudSolutionProvider);
                    }

                    this.PrintUsageAggregatesReportCsp(startColumnNumber, rowNumber, usageAggregates, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.items.Length;

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.links?.self?.uri;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    string content = await BillingUtils.GetUsageAggregates(token, continuationLink);
                    usageAggregates = JsonConvert.DeserializeObject<CspUsageAggregates>(content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                this.FormatTags(UsageApi.CloudSolutionProvider, rowNumber, currentActiveWorksheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to get usage report: {ex.Message}\r\n\r\n\r\n{ex.StackTrace}\r\n",
                    "Get Usage Report (CSP)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private async void GetEaUsageReportButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!this.ValidateUsageReportInput(UsageApi.EnterpriseAgreement))
            {
                return;
            }

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Getting usage report (standard)...";

                var enrollmentNumber = this.EnrollmentNumberComboBox.Text;
                var apiKey = this.EaApiKeyComboBox.Text;
                var reportStartDate = this.StartDateEditBox.Text;
                var reportEndDate = this.EndDateEditBox.Text;

                // Write the report line items:
                int startColumnNumber = 1; // A
                int headerRowNumber = 10;
                int rowNumber = headerRowNumber + 2;
                Excel.Worksheet currentActiveWorksheet = null;
                int currentContinuationCount = 0;
                EaUsageAggregates usageAggregates = await BillingUtils.GetUsageAggregatesEa(apiKey, enrollmentNumber, reportStartDate, reportEndDate);

                do
                {
                    if (usageAggregates == null)
                    {
                        MessageBox.Show(
                            $"ERROR: Failed to get usage report. Verify the correct parameters were provided for Enrollment Number, API Key, Start Date, and End Date and try again.",
                            "Get Usage Report (EA)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (currentContinuationCount == 0)
                    {
                        // Add a fresh worksheet.
                        Excel.Worksheet previousActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
                        currentActiveWorksheet = Globals.ThisAddIn.Application.Worksheets.Add(previousActiveWorksheet);
                        this.PrintUsageAggregatesHeader(startColumnNumber, headerRowNumber, currentActiveWorksheet, UsageApi.EnterpriseAgreement);
                    }

                    this.PrintUsageAggregatesReportEa(startColumnNumber, rowNumber, usageAggregates, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.data.Length;

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.nextLink;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    string content = await BillingUtils.GetUsageAggregates(apiKey, continuationLink);
                    usageAggregates = JsonConvert.DeserializeObject<EaUsageAggregates>(content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                this.FormatTags(UsageApi.EnterpriseAgreement, rowNumber, currentActiveWorksheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to get usage report: {ex.Message}\r\n\r\n\r\n{ex.StackTrace}\r\n",
                    "Get Usage Report (EA)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private bool ValidateUsageReportInput(UsageApi usageApi)
        {
            if (string.IsNullOrWhiteSpace(this.TenantIdComboBox.Text))
            {
                MessageBox.Show($"ERROR: Tenant Id must be specified.", "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.SubscriptionIdComboBox.Text))
            {
                MessageBox.Show($"ERROR: Subscription Id must be specified.", "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.StartDateEditBox.Text))
            {
                MessageBox.Show($"ERROR: Report Start Date (yyyy-mm-dd) must be specified.", "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.EndDateEditBox.Text))
            {
                MessageBox.Show($"ERROR: Report End Date (yyyy-mm-dd) must be specified.", "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (usageApi == UsageApi.EnterpriseAgreement)
            {
                if (string.IsNullOrWhiteSpace(this.EnrollmentNumberComboBox.Text))
                {
                    MessageBox.Show($"ERROR: Enrollment Number must be specified for an EA Usage Report.", "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            SecurityUtils.SaveUsageReportParameters(new PersistedData()
            {
                SubscriptionId = this.SubscriptionIdComboBox.Text.Trim(),
                TenantId = this.TenantIdComboBox.Text.Trim(),
                EnrollmentNumber = this.EnrollmentNumberComboBox.Text.Trim(),
                EaApiKey = this.EaApiKeyComboBox.Text.Trim()
            });

            return true;
        }

        private void FormatTags(UsageApi usageApi, int lastRowNumber, Excel.Worksheet currentActiveWorksheet)
        {
            switch (usageApi)
            {
                case UsageApi.CloudSolutionProvider:
                    currentActiveWorksheet.get_Range($"J11:J{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
                case UsageApi.EnterpriseAgreement:
                    currentActiveWorksheet.get_Range($"Q12:Q{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
                default:
                    currentActiveWorksheet.get_Range($"M11:M{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
            }
        }

        private void PrintUsageAggregatesHeader(int startColumnNumber, int headerRowNumber, Excel.Worksheet currentActiveWorksheet, UsageApi usageApi)
        {
            Globals.ThisAddIn.Application.StatusBar = "Displaying usage report header...";

            var tenantId = this.TenantIdComboBox.Text;
            var subscriptionId = this.SubscriptionIdComboBox.Text;
            var reportStartDate = this.StartDateEditBox.Text;
            var reportEndDate = this.EndDateEditBox.Text;
            var aggregationGranularity = this.AggregationGranularityDropDown.SelectedItem.Tag;
            var tableFirstRowNumber = "11";

            // Write the report header:
            ExcelUtils.WriteHeaderRow("A1", "B1", new[] { "Subscription Id:", $"{subscriptionId}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow("A2", "B2", new[] { "Tenant Id:", $"{tenantId}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow("A3", "B3", new[] { "Report Start Date (UTC):", $"{reportStartDate}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow("A4", "B4", new[] { "Report End Date (UTC):", $"{reportEndDate}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow("A5", "B5", new[] { "Aggregation Granularity:", $"{aggregationGranularity}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow("A6", "B6", new[] { "Report generated (UTC):", $"{DateTime.UtcNow}" }, currentActiveWorksheet);
            if (usageApi == UsageApi.EnterpriseAgreement)
            {
                ExcelUtils.WriteHeaderRow("A7", "B7", new[] { "Enrollment Number:", $"{tenantId}" }, currentActiveWorksheet);
                tableFirstRowNumber = "12";
            }

            ExcelUtils.WriteUsageLineItemHeader(startColumnNumber, headerRowNumber, this.GetHeaderCaptions(usageApi), currentActiveWorksheet);

            // Format the data types of datatime and numeric columns.
            currentActiveWorksheet.get_Range($"A{tableFirstRowNumber}").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // UsageStartTime
            if (usageApi == UsageApi.Standard || usageApi == UsageApi.CloudSolutionProvider)
            {
                currentActiveWorksheet.get_Range($"B{tableFirstRowNumber}").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // UsageEndTime
            }

            //currentActiveWorksheet.get_Range("K11").EntireColumn.NumberFormat = "#####.###################"; // Quantity
        }

        private string[] GetHeaderCaptions(UsageApi usageApi)
        {
            switch (usageApi)
            {
                case UsageApi.CloudSolutionProvider:
                    return this.HeaderCaptionsCsp;
                case UsageApi.EnterpriseAgreement:
                    return this.HeaderCaptionsEa;
                default:
                    return this.HeaderCaptions;
            }
        }

        private void PrintUsageAggregatesReport(int startColumnNumber, int rowNumber, UsageAggregates usageAggregates, int chunkNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying standard usage report chunk {chunkNumber}...";

            foreach (var usageAggregate in usageAggregates.value)
            {
                ExcelUtils.WriteUsageLineItem(startColumnNumber, rowNumber, usageAggregate, this.HeaderCaptions.Length, currentActiveWorksheet);
                rowNumber++;
            }
        }

        private void PrintUsageAggregatesReportCsp(int startColumnNumber, int rowNumber, CspUsageAggregates usageAggregates, int chunkNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying CSP usage report chunk {chunkNumber}...";

            foreach (var usageAggregate in usageAggregates.items)
            {
                ExcelUtils.WriteUsageLineItemCsp(startColumnNumber, rowNumber, usageAggregate, this.HeaderCaptions.Length, currentActiveWorksheet);
                rowNumber++;
            }
        }

        private void PrintUsageAggregatesReportEa(int startColumnNumber, int rowNumber, EaUsageAggregates usageAggregates, int chunkNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying EA usage report chunk {chunkNumber}...";

            foreach (var usageAggregate in usageAggregates.data)
            {
                ExcelUtils.WriteUsageLineItemEa(startColumnNumber, rowNumber, usageAggregate, this.HeaderCaptions.Length, currentActiveWorksheet);
                rowNumber++;
            }
        }

        private void HydrateFromPersistedData()
        {
            var persistedData = SecurityUtils.GetSavedUsageReportParameters();
            if (persistedData != null)
            {
                var ribbonFactory = Globals.Factory.GetRibbonFactory();
                if (!string.IsNullOrWhiteSpace(persistedData.SubscriptionId))
                {
                    this.SubscriptionIdComboBox.Text = persistedData.SubscriptionId;
                    var ribbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = persistedData.SubscriptionId;
                    this.SubscriptionIdComboBox.Items.Add(ribbonDropDownItem);
                }

                if (!string.IsNullOrWhiteSpace(persistedData.TenantId))
                {
                    this.TenantIdComboBox.Text = persistedData.TenantId;
                    var ribbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = persistedData.TenantId;
                    this.TenantIdComboBox.Items.Add(ribbonDropDownItem);
                }

                if (!string.IsNullOrWhiteSpace(persistedData.EnrollmentNumber))
                {
                    this.EnrollmentNumberComboBox.Text = persistedData.EnrollmentNumber;
                    var ribbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = persistedData.EnrollmentNumber;
                    this.EnrollmentNumberComboBox.Items.Add(ribbonDropDownItem);
                }

                if (!string.IsNullOrWhiteSpace(persistedData.EaApiKey))
                {
                    this.EaApiKeyComboBox.Text = persistedData.EaApiKey;
                    var ribbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = persistedData.EaApiKey;
                    this.EaApiKeyComboBox.Items.Add(ribbonDropDownItem);
                }
                this.EaApiKeyComboBox.Text = persistedData.EaApiKey ?? string.Empty;
            }
        }

        private void UpdateAddinButton_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start(AddinInstallUrl);
        }
    }
}
