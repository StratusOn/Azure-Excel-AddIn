﻿using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
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
            "Meter Category", "Meter Sub-Category", "Quantity", "Unit", "Tags", "Resource URI", "Location", "Order Number", "Part Number", "Additional Info",
            "Metered Region", "Metered Service", "Metered Service Type", "Project", "Service Info1"
        };
        private readonly string[] HeaderCaptionsCsp = {
            "Usage Start Time (UTC)", "Usage End Time (UTC)", "Meter Id", "Meter Name", "Meter Category", "Meter Sub-Category", "Meter Region",
            "Quantity", "Unit", "Tags", "Resource URI", "Location", "Order Number", "Part Number", "Additional Info",
            "Metered Region", "Metered Service", "Metered Service Type", "Project", "Service Info1", "Attributes"
        };
        private readonly string[] HeaderCaptionsEa = {
            "Usage Time (UTC)", "Account Id", "Account Name", "Product Id", "Product", "Resource Location Id", "Resource Location", "Consumed Service Id", "Consumed Service", "Department Id", "Department Name",
            "Account Owner Email", "Service Administrator Id", "Subscription Id", "Subscription Guid", "Subscription Name", "Tags",
            "Meter Id", "Meter Name", "Meter Category", "Meter Sub-Category", "Meter Region", "Consumed Quantity", "Unit of Measure", "Resource Rate", "Cost",
            "Instance Id", "Service Info 1", "Service Info 2", "Additional Info", "Store Service Identifier", "Cost Center", "Resource Group"
        };

        private readonly string[] HeaderCaptionRateCard =
        {
            "Currency", "Locale", "Meter Region", "Is Tax Included", "Tags"
        };
        private readonly string[] HeaderCaptionRateCardOfferTerm =
        {
            "Name", "Credit", "Effective Date", "Excluded Meter Ids", "Tiered Discount", "Initial Discount"
        };
        private readonly string[] HeaderCaptionsRateCardMeter =
        {
            "Meter Id", "Meter Name", "Meter Category", "Meter Sub-Category", "Unit", "Meter Region", "Meter Rates", "Initial Rate", "Meter Tags", "Effective Date", "Included Quantity", "Meter Status"
        };

        private void AzureRibbonTab_Load(object sender, RibbonUIEventArgs e)
        {
            var today = DateTime.Today;
            var yesterday = today.AddDays(-1);
            this.StartDateEditBox.Text = $"{yesterday.Year}-{yesterday.Month:0#}-{yesterday.Day:0#}";
            this.EndDateEditBox.Text = $"{today.Year}-{today.Month:0#}-{today.Day:0#}";

            this.HydrateFromPersistedData();
            this.SetControlsEnableState();
        }

        private void GetTokenButton_Click(object sender, RibbonControlEventArgs e)
        {
            var tenantId = this.AuthTenantIdEditBox.Text.Trim();
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
                string token = AuthUtils.GetAuthorizationHeader(tenantId, true, usageApi, (AzureEnvironment)Enum.Parse(typeof(AzureEnvironment), (string)this.AzureEnvironmentDropDown.SelectedItem.Tag));

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
            string subscriptionType = this.SubscriptionTypeDropDown.SelectedItem.Tag as string;
            string reportType = this.ReportTypeDropDown.SelectedItem.Tag as string;
            switch (subscriptionType)
            {
                case "CSP":
                    if (reportType == "Usage")
                    {
                        await this.GetCspUsageReportAsync();
                    }
                    else
                    {
                        await this.GetCspRateCardAsync();
                    }
                    break;
                case "EA":
                    if (reportType == "Usage")
                    {
                        await this.GetEaUsageReportAsync();
                    }
                    else
                    {
                        await this.GetEaRateCardAsync();
                    }
                    break;
                default:
                    if (reportType == "Usage")
                    {
                        await this.GetStandardUsageReportAsync();
                    }
                    else
                    {
                        await this.GetStandardRateCardAsync();
                    }
                    break;
            }
        }

        private void ReportTypeDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            this.SetControlsEnableState();
        }

        private void SubscriptionTypeDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            this.SetControlsEnableState();
        }

        private void UpdateAddinButton_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start(AddinInstallUrl);
        }

        private async Task GetStandardUsageReportAsync()
        {
            if (!this.ValidateUsageReportInput(UsageApi.Standard))
            {
                return;
            }

            var tenantId = this.TenantIdComboBox.Text.Trim();

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Authenticating...";

                string token = AuthUtils.GetAuthorizationHeader(
                    tenantId, 
                    this.ForceReAuthCheckBox.Checked, 
                    UsageApi.Standard, 
                    this.ApplicationIdComboBox.Text.Trim(), 
                    this.AppKeyComboBox.Text.Trim(), 
                    (AzureEnvironment)Enum.Parse(typeof(AzureEnvironment), (string)this.AzureEnvironmentDropDown.SelectedItem.Tag));

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show(
                        $"ERROR: Failed to acquire a token. Verify you entered the right credentials and the correct Tenant Id, and make sure 'Force Re-Authentication' is checked and try again.",
                        "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Globals.ThisAddIn.Application.StatusBar = "Getting usage report (standard)...";

                var subscriptionId = this.SubscriptionIdComboBox.Text.Trim();
                var reportStartDate = this.StartDateEditBox.Text.Trim();
                var reportEndDate = this.EndDateEditBox.Text.Trim();
                var aggregationGranularity = (string)this.AggregationGranularityDropDown.SelectedItem.Tag;
                var showDetails = "true"; // this.ShowDetailsCheckBox.Checked.ToString();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int startHeaderRowNumber = 1;
                int rowNumber = 0;
                int currentContinuationCount = 0;
                Excel.Worksheet currentActiveWorksheet = null;

                UsageAggregates usageAggregates = await BillingUtils.GetUsageAggregatesStandardAsync(token, subscriptionId, reportStartDate, reportEndDate, aggregationGranularity, showDetails);

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
                        currentActiveWorksheet.SetWorksheetName(UsageApi.Standard, BillingApiType.Usage);
                        rowNumber = this.PrintUsageAggregatesHeader(startColumnNumber, startHeaderRowNumber, currentActiveWorksheet, UsageApi.Standard);
                    }

                    this.PrintUsageAggregatesReport(startColumnNumber, rowNumber, usageAggregates, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.value.Count;

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.nextLink;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    string content = await BillingUtils.GetRestCallResultsAsync(token, continuationLink);
                    usageAggregates = JsonConvert.DeserializeObject<UsageAggregates>(content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                //this.FormatTags(UsageApi.Standard, rowNumber, currentActiveWorksheet);
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

        private async Task GetCspUsageReportAsync()
        {
            if (!this.ValidateUsageReportInput(UsageApi.CloudSolutionProvider))
            {
                return;
            }

            var tenantId = this.TenantIdComboBox.Text.Trim();

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Authenticating...";

                string token = AuthUtils.GetAuthorizationHeader(
                    tenantId, 
                    this.ForceReAuthCheckBox.Checked, 
                    UsageApi.CloudSolutionProvider, 
                    this.ApplicationIdComboBox.Text.Trim(), 
                    this.AppKeyComboBox.Text.Trim(),
                    (AzureEnvironment)Enum.Parse(typeof(AzureEnvironment), (string)this.AzureEnvironmentDropDown.SelectedItem.Tag));

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show(
                        $"ERROR: Failed to acquire a token. Verify you entered the right credentials and the correct Tenant Id, and make sure 'Force Re-Authentication' is checked and try again.",
                        "Get Usage Report (CSP)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Globals.ThisAddIn.Application.StatusBar = "Getting usage report (CSP)...";

                var subscriptionId = this.SubscriptionIdComboBox.Text.Trim();
                var reportStartDate = this.StartDateEditBox.Text.Trim();
                var reportEndDate = this.EndDateEditBox.Text.Trim();
                var aggregationGranularity = (string)this.AggregationGranularityDropDown.SelectedItem.Tag;
                var showDetails = "true"; // this.ShowDetailsCheckBox.Checked.ToString();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int startHeaderRowNumber = 1;
                int rowNumber = 0;
                int currentContinuationCount = 0;
                int chunkSize = DefaultChunkSize;
                Excel.Worksheet currentActiveWorksheet = null;

                CspUsageAggregates usageAggregates = await BillingUtils.GetUsageAggregatesCspAsync(token, subscriptionId, tenantId, reportStartDate, reportEndDate, aggregationGranularity, showDetails, chunkSize);

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
                        currentActiveWorksheet.SetWorksheetName(UsageApi.CloudSolutionProvider, BillingApiType.Usage);
                        rowNumber = this.PrintUsageAggregatesHeader(startColumnNumber, startHeaderRowNumber, currentActiveWorksheet, UsageApi.CloudSolutionProvider);
                    }

                    this.PrintUsageAggregatesReportCsp(startColumnNumber, rowNumber, usageAggregates, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.items.Count;

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.links?.self?.uri;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    string content = await BillingUtils.GetRestCallResultsAsync(token, continuationLink);
                    usageAggregates = JsonConvert.DeserializeObject<CspUsageAggregates>(content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                //this.FormatTags(UsageApi.CloudSolutionProvider, rowNumber, currentActiveWorksheet);
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

        private async Task GetEaUsageReportAsync()
        {
            if (!this.ValidateUsageReportInput(UsageApi.EnterpriseAgreement))
            {
                return;
            }

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Getting usage report (EA)...";

                var enrollmentNumber = this.EnrollmentNumberComboBox.Text.Trim();
                var apiKey = this.EaApiKeyComboBox.Text.Trim();
                var reportStartDate = this.StartDateEditBox.Text.Trim();
                var reportEndDate = this.EndDateEditBox.Text.Trim();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int startHeaderRowNumber = 1;
                int rowNumber = 0;
                int currentContinuationCount = 0;
                Excel.Worksheet currentActiveWorksheet = null;

                EaUsageAggregates usageAggregates = await BillingUtils.GetUsageAggregatesEaAsync(apiKey, enrollmentNumber, reportStartDate, reportEndDate);

                do
                {
                    if (usageAggregates == null)
                    {
                        MessageBox.Show(
                            $"ERROR: Failed to get usage report. Verify the correct parameters were provided for Enrollment Number, API Key, Start Date, and End Date (or Billing Period) and try again.",
                            "Get Usage Report (EA)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (currentContinuationCount == 0)
                    {
                        // Add a fresh worksheet.
                        Excel.Worksheet previousActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
                        currentActiveWorksheet = Globals.ThisAddIn.Application.Worksheets.Add(previousActiveWorksheet);
                        currentActiveWorksheet.SetWorksheetName(UsageApi.EnterpriseAgreement, BillingApiType.Usage);
                        rowNumber = this.PrintUsageAggregatesHeader(startColumnNumber, startHeaderRowNumber, currentActiveWorksheet, UsageApi.EnterpriseAgreement);
                    }

                    this.PrintUsageAggregatesReportEa(startColumnNumber, rowNumber, usageAggregates, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.data.Count;

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.nextLink;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    string content = await BillingUtils.GetRestCallResultsAsync(apiKey, continuationLink);
                    usageAggregates = JsonConvert.DeserializeObject<EaUsageAggregates>(content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                //this.FormatTags(UsageApi.EnterpriseAgreement, rowNumber, currentActiveWorksheet);
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

        private async Task GetStandardRateCardAsync()
        {
            if (!this.ValidateUsageReportInput(UsageApi.Standard))
            {
                return;
            }

            var tenantId = this.TenantIdComboBox.Text.Trim();

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Authenticating...";

                string token = AuthUtils.GetAuthorizationHeader(
                    tenantId, 
                    this.ForceReAuthCheckBox.Checked, 
                    UsageApi.Standard, 
                    null, 
                    null, 
                    (AzureEnvironment)Enum.Parse(typeof(AzureEnvironment), (string)this.AzureEnvironmentDropDown.SelectedItem.Tag));

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show(
                        $"ERROR: Failed to acquire a token. Verify you entered the right credentials and the correct Tenant Id, and make sure 'Force Re-Authentication' is checked and try again.",
                        "Get Rate Card", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Globals.ThisAddIn.Application.StatusBar = "Getting rate card (standard)...";

                var subscriptionId = this.SubscriptionIdComboBox.Text.Trim();
                var offerDurableId = this.RateCardOfferDurableIdComboBox.Text.Trim();
                var currency = this.RateCardCurrencyComboBox.Text.Trim();
                var locale = this.RateCardLocaleComboBox.Text.Trim();
                var regionInfo = this.RateCardRegionInfoComboBox.Text.Trim();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int startHeaderRowNumber = 1;

                RateCard rateCard = await BillingUtils.GetRateCardStandardAsync(token, subscriptionId, offerDurableId, currency, locale, regionInfo);
                if (rateCard == null)
                {
                    MessageBox.Show(
                        $"ERROR: Failed to get the rate card. Verify the correct parameters were provided for Subscription Id, Offer Id, currency, locale, and region info and try again.",
                        "Get Rate Card", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Add a fresh worksheet and write the results.
                Excel.Worksheet currentActiveWorksheet = Globals.ThisAddIn.Application.Worksheets.Add(Globals.ThisAddIn.Application.ActiveSheet);
                currentActiveWorksheet.SetWorksheetName(UsageApi.Standard, BillingApiType.RateCard);
                var rowNumber = this.PrintRateCardHeader(startColumnNumber, startHeaderRowNumber, rateCard, currentActiveWorksheet, UsageApi.Standard);
                this.PrintRateCardReport(startColumnNumber, rowNumber, rateCard, currentActiveWorksheet);
                rowNumber += rateCard.Meters.Count;
                //this.FormatRate(UsageApi.Standard, rowNumber, currentActiveWorksheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to get rate card: {ex.Message}\r\n\r\n\r\n{ex.StackTrace}\r\n",
                    "Get Rate Card", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private async Task GetCspRateCardAsync()
        {
            MessageBox.Show("Not Implemented");
        }

        private async Task GetEaRateCardAsync()
        {
            MessageBox.Show("Not Implemented");
        }

        private void FormatTags(UsageApi usageApi, int lastRowNumber, Excel.Worksheet currentActiveWorksheet)
        {
            switch (usageApi)
            {
                case UsageApi.CloudSolutionProvider:
                    currentActiveWorksheet.get_Range($"J11:J{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
                case UsageApi.EnterpriseAgreement:
                    currentActiveWorksheet.get_Range($"Q11:Q{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
                default:
                    currentActiveWorksheet.get_Range($"M11:M{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
            }
        }

        private void FormatRate(UsageApi usageApi, int lastRowNumber, Excel.Worksheet currentActiveWorksheet)
        {
            switch (usageApi)
            {
                case UsageApi.CloudSolutionProvider:
                    currentActiveWorksheet.get_Range($"G11:G{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
                case UsageApi.EnterpriseAgreement:
                    currentActiveWorksheet.get_Range($"G11:G{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
                default:
                    currentActiveWorksheet.get_Range($"G11:G{lastRowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC
                    break;
            }
        }

        private int PrintUsageAggregatesHeader(int startColumnNumber, int headerRowNumber, Excel.Worksheet currentActiveWorksheet, UsageApi usageApi)
        {
            Globals.ThisAddIn.Application.StatusBar = "Displaying usage report header...";

            var tenantId = this.TenantIdComboBox.Text;
            var subscriptionId = this.SubscriptionIdComboBox.Text;
            var reportStartDate = this.StartDateEditBox.Text;
            var reportEndDate = this.EndDateEditBox.Text;
            var aggregationGranularity = this.AggregationGranularityDropDown.SelectedItem.Tag;

            // Write the report header:
            int rowNumber = headerRowNumber;
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Subscription Id:", $"{subscriptionId}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Tenant Id:", $"{tenantId}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Report Start Date (UTC):", $"{reportStartDate}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Report End Date (UTC):", $"{reportEndDate}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Aggregation Granularity:", $"{aggregationGranularity}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Report generated (UTC):", $"{DateTime.UtcNow}" }, currentActiveWorksheet);
            if (usageApi == UsageApi.EnterpriseAgreement)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Enrollment Number:", $"{tenantId}" }, currentActiveWorksheet);
            }

            ExcelUtils.WriteUsageLineItemHeader(startColumnNumber, rowNumber, this.GetHeaderCaptions(usageApi), currentActiveWorksheet);

            // Format the data types of datatime and numeric columns.
            rowNumber++; // Starting row number for the table.
            currentActiveWorksheet.get_Range($"A{rowNumber}").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // UsageStartTime
            if (usageApi == UsageApi.Standard || usageApi == UsageApi.CloudSolutionProvider)
            {
                currentActiveWorksheet.get_Range($"B{rowNumber}").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // UsageEndTime
            }

            return ++rowNumber; // return the first writable row number.
        }

        private int PrintRateCardHeader(int startColumnNumber, int headerRowNumber, RateCard rateCard, Excel.Worksheet currentActiveWorksheet, UsageApi usageApi)
        {
            Globals.ThisAddIn.Application.StatusBar = "Displaying rate card header...";

            var tenantId = this.TenantIdComboBox.Text.Trim();
            var subscriptionId = this.SubscriptionIdComboBox.Text.Trim();
            var offerDurableId = this.RateCardOfferDurableIdComboBox.Text.Trim();
            var currency = this.RateCardCurrencyComboBox.Text.Trim();
            var locale = this.RateCardLocaleComboBox.Text.Trim();
            var regionInfo = this.RateCardRegionInfoComboBox.Text.Trim();

            // Write the report header:
            int rowNumber = headerRowNumber;
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Subscription Id:", $"{subscriptionId}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Tenant Id:", $"{tenantId}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Durable Id:", $"{offerDurableId}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Currency:", $"{currency}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Locale:", $"{locale}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Region Info:", $"{regionInfo}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Is Tax Included:", $"{rateCard.IsTaxIncluded}" }, currentActiveWorksheet);
            if (rateCard.Tags != null)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "RateCard Tags:", $"{string.Join(";", rateCard.Tags)}" }, currentActiveWorksheet);
            }
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "RateCard generated (UTC):", $"{DateTime.UtcNow}" }, currentActiveWorksheet);
            if (rateCard.OfferTerms.Count > 0)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Terms Name:", $"{rateCard.OfferTerms[0].Name}" }, currentActiveWorksheet);
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Terms Duration:", $"{rateCard.OfferTerms[0].EffectiveDate}" }, currentActiveWorksheet);
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Credit:", $"{rateCard.OfferTerms[0].Credit}" }, currentActiveWorksheet);
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Excluded Meter Ids:", $"{string.Join(";", rateCard.OfferTerms[0].ExcludedMeterIds)}" }, currentActiveWorksheet);
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Terms Duration:", $"{string.Join(";", rateCard.OfferTerms[0].TieredDiscount.Select(x => x.Key + "=" + x.Value))}" }, currentActiveWorksheet);
            }

            ExcelUtils.WriteUsageLineItemHeader(startColumnNumber, rowNumber, this.GetRateCardHeaderCaptions(usageApi), currentActiveWorksheet);

            // Format the data types of datatime and numeric columns.
            rowNumber++; // Starting row number for the table.
            currentActiveWorksheet.get_Range($"J{rowNumber}").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // EffectiveDate

            return ++rowNumber; // return the first writable row number.
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

        private string[] GetRateCardHeaderCaptions(UsageApi usageApi)
        {
            switch (usageApi)
            {
                case UsageApi.CloudSolutionProvider:
                    return this.HeaderCaptionsRateCardMeter;
                case UsageApi.EnterpriseAgreement:
                    return this.HeaderCaptionsRateCardMeter;
                default:
                    return this.HeaderCaptionsRateCardMeter;
            }
        }

        private void PrintUsageAggregatesReport(int startColumnNumber, int rowNumber, UsageAggregates usageAggregates, int chunkNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying standard usage report chunk {chunkNumber}. Please wait...";

            if (usageAggregates.value != null)
            {
                foreach (var usageAggregate in usageAggregates.value)
                {
                    ExcelUtils.WriteUsageLineItem(startColumnNumber, rowNumber, usageAggregate,
                        this.HeaderCaptions.Length, currentActiveWorksheet);
                    rowNumber++;
                }
            }
        }

        private void PrintUsageAggregatesReportCsp(int startColumnNumber, int rowNumber, CspUsageAggregates usageAggregates, int chunkNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying CSP usage report chunk {chunkNumber}. Please wait...";

            if (usageAggregates.items != null)
            {
                foreach (var usageAggregate in usageAggregates.items)
                {
                    ExcelUtils.WriteUsageLineItemCsp(startColumnNumber, rowNumber, usageAggregate,
                        this.HeaderCaptionsCsp.Length, currentActiveWorksheet);
                    rowNumber++;
                }
            }
        }

        private void PrintUsageAggregatesReportEa(int startColumnNumber, int rowNumber, EaUsageAggregates usageAggregates, int chunkNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying EA usage report chunk {chunkNumber}. Please wait...";

            if (usageAggregates.data != null)
            {
                foreach (var usageAggregate in usageAggregates.data)
                {
                    ExcelUtils.WriteUsageLineItemEa(startColumnNumber, rowNumber, usageAggregate,
                        this.HeaderCaptionsEa.Length, currentActiveWorksheet);
                    rowNumber++;
                }
            }
        }

        private void PrintRateCardReport(int startColumnNumber, int rowNumber, RateCard rateCard, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying standard rate card. Please wait...";

            if (rateCard.Meters != null)
            {
                foreach (var meter in rateCard.Meters)
                {
                    ExcelUtils.WriteRateCardMeterLineItem(startColumnNumber, rowNumber, meter,
                        this.HeaderCaptionsRateCardMeter.Length, currentActiveWorksheet);
                    rowNumber++;
                }
            }
        }

        private void PersistData()
        {
            SecurityUtils.SaveUsageReportParameters(new PersistedData()
            {
                SubscriptionId = this.SubscriptionIdComboBox.Text.Trim(),
                TenantId = this.TenantIdComboBox.Text.Trim(),
                EnrollmentNumber = this.EnrollmentNumberComboBox.Text.Trim(),
                EaApiKey = this.EaApiKeyComboBox.Text.Trim(),
                ApplicationId = this.ApplicationIdComboBox.Text.Trim(),
                ApplicationKey = this.AppKeyComboBox.Text.Trim()
            });
        }

        private void AddDataToCombos()
        {
            var ribbonFactory = Globals.Factory.GetRibbonFactory();

            var subscriptionIdRibbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
            subscriptionIdRibbonDropDownItem.Label = this.SubscriptionIdComboBox.Text.Trim();
            if (this.SubscriptionIdComboBox.Items.All(item => string.Compare(item.Label, subscriptionIdRibbonDropDownItem.Label, StringComparison.CurrentCultureIgnoreCase) != 0))
            {
                this.SubscriptionIdComboBox.Items.Add(subscriptionIdRibbonDropDownItem);
            }

            var tenantIdRibbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
            tenantIdRibbonDropDownItem.Label = this.TenantIdComboBox.Text;
            if (this.TenantIdComboBox.Items.All(item => string.Compare(item.Label, tenantIdRibbonDropDownItem.Label, StringComparison.CurrentCultureIgnoreCase) != 0))
            {
                this.TenantIdComboBox.Items.Add(tenantIdRibbonDropDownItem);
            }

            var applicationIdRibbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
            applicationIdRibbonDropDownItem.Label = this.ApplicationIdComboBox.Text;
            if (this.ApplicationIdComboBox.Items.All(item => string.Compare(item.Label, applicationIdRibbonDropDownItem.Label, StringComparison.CurrentCultureIgnoreCase) != 0))
            {
                this.ApplicationIdComboBox.Items.Add(applicationIdRibbonDropDownItem);
            }

            var appKeyRibbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
            appKeyRibbonDropDownItem.Label = this.AppKeyComboBox.Text;
            if (this.AppKeyComboBox.Items.All(item => string.Compare(item.Label, appKeyRibbonDropDownItem.Label, StringComparison.CurrentCultureIgnoreCase) != 0))
            {
                this.AppKeyComboBox.Items.Add(appKeyRibbonDropDownItem);
            }

            var enrollmentNumberRibbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
            enrollmentNumberRibbonDropDownItem.Label = this.EnrollmentNumberComboBox.Text;
            if (this.EnrollmentNumberComboBox.Items.All(item => string.Compare(item.Label, enrollmentNumberRibbonDropDownItem.Label, StringComparison.CurrentCultureIgnoreCase) != 0))
            {
                this.EnrollmentNumberComboBox.Items.Add(enrollmentNumberRibbonDropDownItem);
            }

            var eaApiKeyRibbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
            eaApiKeyRibbonDropDownItem.Label = this.EaApiKeyComboBox.Text;
            if (this.EaApiKeyComboBox.Items.All(item => string.Compare(item.Label, eaApiKeyRibbonDropDownItem.Label, StringComparison.CurrentCultureIgnoreCase) != 0))
            {
                this.EaApiKeyComboBox.Items.Add(eaApiKeyRibbonDropDownItem);
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

                if (!string.IsNullOrWhiteSpace(persistedData.ApplicationId))
                {
                    this.ApplicationIdComboBox.Text = persistedData.ApplicationId;
                    var ribbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = persistedData.ApplicationId;
                    this.ApplicationIdComboBox.Items.Add(ribbonDropDownItem);
                }

                if (!string.IsNullOrWhiteSpace(persistedData.ApplicationKey))
                {
                    this.AppKeyComboBox.Text = persistedData.ApplicationKey;
                    var ribbonDropDownItem = ribbonFactory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = persistedData.ApplicationKey;
                    this.AppKeyComboBox.Items.Add(ribbonDropDownItem);
                }
            }
        }

        private void SetControlsEnableState()
        {
            string reportType = this.ReportTypeDropDown.SelectedItem.Tag as string;
            string subscriptionType = this.SubscriptionTypeDropDown.SelectedItem.Tag as string;
            bool isUsageReport = reportType == "Usage";
            bool isRateCard = reportType == "RateCard";
            bool isStandard = subscriptionType == "Standard";
            bool isCsp = subscriptionType == "CSP";
            bool isEa = subscriptionType == "EA";

            this.RateCardOfferDurableIdComboBox.Enabled = isRateCard && isStandard;
            this.RateCardCurrencyComboBox.Enabled = isRateCard && (isStandard || isCsp);
            this.RateCardLocaleComboBox.Enabled = isRateCard && (isStandard || isCsp);
            this.RateCardRegionInfoComboBox.Enabled = isRateCard && (isStandard || isCsp);
            this.PriceSheetBillingPeriodComboBox.Enabled = isEa;
            this.EaApiKeyComboBox.Enabled = isEa;
            this.ApplicationIdComboBox.Enabled = isStandard || isCsp;
            this.AppKeyComboBox.Enabled = isStandard || isCsp;
            this.EnrollmentNumberComboBox.Enabled = isEa;
            this.AggregationGranularityDropDown.Enabled = isUsageReport && (isStandard || isCsp);
            this.StartDateEditBox.Enabled = isUsageReport;
            this.EndDateEditBox.Enabled = isUsageReport;
            this.SubscriptionIdComboBox.Enabled = isStandard || isCsp;
            this.TenantIdComboBox.Enabled = true;
        }

        private bool ValidateUsageReportInput(UsageApi usageApi)
        {
            if (string.IsNullOrWhiteSpace(this.TenantIdComboBox.Text))
            {
                MessageBox.Show($"ERROR: Tenant Id must be specified.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.SubscriptionIdComboBox.Text) && this.SubscriptionIdComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Subscription Id must be specified.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.StartDateEditBox.Text) && this.StartDateEditBox.Enabled)
            {
                MessageBox.Show($"ERROR: Report Start Date (yyyy-mm-dd) must be specified.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.EndDateEditBox.Text) && this.EndDateEditBox.Enabled)
            {
                MessageBox.Show($"ERROR: Report End Date (yyyy-mm-dd) must be specified.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.EnrollmentNumberComboBox.Text) && this.EnrollmentNumberComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Enrollment Number must be specified for an EA Usage Report.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.EaApiKeyComboBox.Text) && this.EaApiKeyComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: An API Key generated from the EA Portal must be specified for an EA Usage Report.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.ApplicationIdComboBox.Text) && !string.IsNullOrWhiteSpace(this.AppKeyComboBox.Text) && this.ApplicationIdComboBox.Enabled && this.AppKeyComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Application Id must be specified when an Application Key is specified.", "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.RateCardLocaleComboBox.Text) && this.RateCardLocaleComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Locale must be specified (e.g. en-US).", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.RateCardCurrencyComboBox.Text) && this.RateCardCurrencyComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Currency must be specified (e.g. USD).", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.RateCardOfferDurableIdComboBox.Text) && this.RateCardOfferDurableIdComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Offer must be specified (e.g. MS-AZR-0003P for Pay-As-You-Go).", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.RateCardRegionInfoComboBox.Text) && this.RateCardRegionInfoComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Region must be specified (e.g.: US).", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            this.PersistData();
            this.AddDataToCombos();

            return true;
        }
    }
}
