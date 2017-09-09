using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
        private readonly string[] HeaderCaptionsRateCardMeter =
        {
            "Meter Id", "Meter Name", "Meter Category", "Meter Sub-Category", "Unit", "Meter Region", "Meter Rates", "Initial Rate", "Meter Tags", "Effective Date", "Included Quantity", "Meter Status"
        };
        private readonly string[] HeaderCaptionsPriceSheetMeter =
        {
            "Id", "Billing Period Id", "Meter Id", "Meter Name", "Meter Region", "Unit of Measure", "Included Quantity", "Part Number", "Unit Price", "Currency Code"
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
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

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
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

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
                        await this.GetRateCardAsync(UsageApi.CloudSolutionProvider);
                    }
                    break;
                case "EA":
                    if (reportType == "Usage")
                    {
                        await this.GetEaUsageReportAsync();
                    }
                    else
                    {
                        await this.GetPriceSheetAsync();
                    }
                    break;
                default:
                    if (reportType == "Usage")
                    {
                        await this.GetStandardUsageReportAsync();
                    }
                    else
                    {
                        await this.GetRateCardAsync(UsageApi.Standard);
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
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            if (!this.ValidateUsageReportInput(UsageApi.Standard))
            {
                return;
            }

            var tenantId = this.TenantIdComboBox.Text.Trim();
            bool includeRawPayload = this.IncludeRawPayloadCheckBox.Checked;
            StringBuilder payload = includeRawPayload ? new StringBuilder() : null;
            string currentChunkContent = null;

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

                var usageAggregates = await BillingUtils.GetUsageAggregatesStandardAsync(token, subscriptionId, reportStartDate, reportEndDate, aggregationGranularity, showDetails);

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

                        if (includeRawPayload)
                        {
                            payload.Append("{\"usage\":[");
                        }
                    }

                    currentChunkContent = usageAggregates.Item2;
                    this.PrintUsageAggregatesReport(startColumnNumber, rowNumber, usageAggregates.Item1.Value, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.Item1.Value.value.Count;

                    if (includeRawPayload)
                    {
                        payload.Append(currentChunkContent);
                    }

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.Item1.Value.nextLink;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    if (includeRawPayload)
                    {
                        payload.Append(",");
                    }

                    string content = await BillingUtils.GetRestCallResultsAsync(token, continuationLink);
                    usageAggregates = new Tuple<Lazy<UsageAggregates>, string>(new Lazy<UsageAggregates>(() => { return JsonConvert.DeserializeObject<UsageAggregates>(content); }), content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                // Add a worksheet for raw payload.
                if (includeRawPayload)
                {
                    payload.Append("]}");
                    ShowRawPayload("usage-std-", FormatJson(payload.ToString()));
                }
            }
            catch (Exception ex)
            {
                if (includeRawPayload && payload.Length > 0)
                {
                    payload.Append(currentChunkContent ?? string.Empty);
                    payload.Append("]}");
                    ShowRawPayload("usage-std-", FormatJson(payload.ToString()));
                }

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
            bool includeRawPayload = this.IncludeRawPayloadCheckBox.Checked;
            StringBuilder payload = includeRawPayload ? new StringBuilder() : null;
            string currentChunkContent = null;

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

                var usageAggregates = await BillingUtils.GetUsageAggregatesCspAsync(token, subscriptionId, tenantId, reportStartDate, reportEndDate, aggregationGranularity, showDetails, chunkSize);

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

                        if (includeRawPayload)
                        {
                            payload.Append("{\"usage\":[");
                        }
                    }

                    currentChunkContent = usageAggregates.Item2;
                    this.PrintUsageAggregatesReportCsp(startColumnNumber, rowNumber, usageAggregates.Item1.Value, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.Item1.Value.items.Count;

                    if (includeRawPayload)
                    {
                        payload.Append(currentChunkContent);
                    }

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.Item1.Value.links?.self?.uri;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    if (includeRawPayload)
                    {
                        payload.Append(",");
                    }

                    string content = await BillingUtils.GetRestCallResultsAsync(token, continuationLink);
                    usageAggregates = new Tuple<Lazy<CspUsageAggregates>, string>(new Lazy<CspUsageAggregates>(() => { return JsonConvert.DeserializeObject<CspUsageAggregates>(content); }), content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                // Add a worksheet for raw payload.
                if (includeRawPayload)
                {
                    payload.Append("]}");
                    ShowRawPayload("usage-csp-", FormatJson(payload.ToString()));
                }
            }
            catch (Exception ex)
            {
                if (includeRawPayload && payload.Length > 0)
                {
                    payload.Append(currentChunkContent ?? string.Empty);
                    payload.Append("]}");
                    ShowRawPayload("usage-csp-", FormatJson(payload.ToString()));
                }

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

            bool includeRawPayload = this.IncludeRawPayloadCheckBox.Checked;
            StringBuilder payload = includeRawPayload ? new StringBuilder() : null;
            string currentChunkContent = null;

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Getting usage report (EA)...";

                var enrollmentNumber = this.EnrollmentNumberComboBox.Text.Trim();
                var apiKey = this.EaApiKeyEditBox.Text.Trim();
                var reportStartDate = this.StartDateEditBox.Text.Trim();
                var reportEndDate = this.EndDateEditBox.Text.Trim();
                var billingPeriod = this.PriceSheetBillingPeriodComboBox.Text.Trim();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int startHeaderRowNumber = 1;
                int rowNumber = 0;
                int currentContinuationCount = 0;
                Excel.Worksheet currentActiveWorksheet = null;

                var usageAggregates = await BillingUtils.GetUsageAggregatesEaAsync(apiKey, enrollmentNumber, reportStartDate, reportEndDate, billingPeriod);

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

                        if (includeRawPayload)
                        {
                            payload.Append("{\"usage\":[");
                        }
                    }

                    currentChunkContent = usageAggregates.Item2;
                    this.PrintUsageAggregatesReportEa(startColumnNumber, rowNumber, usageAggregates.Item1.Value, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.Item1.Value.data.Count;

                    if (includeRawPayload)
                    {
                        payload.Append(currentChunkContent);
                    }

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.Item1.Value.nextLink;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    if (includeRawPayload)
                    {
                        payload.Append(",");
                    }

                    string content = await BillingUtils.GetRestCallResultsAsync(apiKey, continuationLink);
                    usageAggregates = new Tuple<Lazy<EaUsageAggregates>, string>(new Lazy<EaUsageAggregates>(() => { return JsonConvert.DeserializeObject<EaUsageAggregates>(content); }), content);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                // Add a worksheet for raw payload.
                if (includeRawPayload)
                {
                    payload.Append("]}");
                    ShowRawPayload("usage-ea-", FormatJson(payload.ToString()));
                }
            }
            catch (Exception ex)
            {
                // Add a worksheet for raw payload.
                if (includeRawPayload && payload.Length > 0)
                {
                    payload.Append(currentChunkContent ?? string.Empty);
                    payload.Append("]}");
                    ShowRawPayload("usage-ea-", FormatJson(payload.ToString()));
                }

                MessageBox.Show($"ERROR: Failed to get usage report: {ex.Message}\r\n\r\n\r\n{ex.StackTrace}\r\n",
                    "Get Usage Report (EA)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private async Task GetRateCardAsync(UsageApi usageApi)
        {
            if (!this.ValidateUsageReportInput(usageApi))
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
                    usageApi,
                    this.ApplicationIdComboBox.Text.Trim(),
                    this.AppKeyComboBox.Text.Trim(),
                    (AzureEnvironment)Enum.Parse(typeof(AzureEnvironment), (string)this.AzureEnvironmentDropDown.SelectedItem.Tag));

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show(
                        $"ERROR: Failed to acquire a token. Verify you entered the right credentials and the correct Tenant Id, and make sure 'Force Re-Authentication' is checked and try again.",
                        $"Get Rate Card ({usageApi})", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Globals.ThisAddIn.Application.StatusBar = $"Getting rate card ({usageApi})...";

                var subscriptionId = this.SubscriptionIdComboBox.Text.Trim();
                var offerDurableId = this.RateCardOfferDurableIdComboBox.Text.Trim();
                var currency = this.RateCardCurrencyComboBox.Text.Trim();
                var locale = this.RateCardLocaleComboBox.Text.Trim();
                var regionInfo = this.RateCardRegionInfoComboBox.Text.Trim();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int startHeaderRowNumber = 1;

                if (usageApi == UsageApi.CloudSolutionProvider)
                {
                    var rateCard = await BillingUtils.GetRateCardCspAsync(token, currency, locale, regionInfo);
                    if (rateCard == null)
                    {
                        MessageBox.Show(
                            $"ERROR: Failed to get the rate card. Verify the correct parameters were provided for currency, locale, and region info and try again.",
                            $"Get Rate Card ({usageApi})", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Add a worksheet for raw payload.
                    if (this.IncludeRawPayloadCheckBox.Checked && !string.IsNullOrWhiteSpace(rateCard.Item2))
                    {
                        ShowRawPayload("ratecard-csp-", FormatJson(rateCard.Item2));
                    }

                    // Add a fresh worksheet and write the results.
                    Excel.Worksheet currentActiveWorksheet =
                        Globals.ThisAddIn.Application.Worksheets.Add(Globals.ThisAddIn.Application.ActiveSheet);
                    currentActiveWorksheet.SetWorksheetName(usageApi, BillingApiType.RateCard);
                    var rowNumber = this.PrintCspRateCardHeader(startColumnNumber, startHeaderRowNumber, rateCard.Item1.Value,
                        currentActiveWorksheet, usageApi);
                    this.PrintCspRateCardReport(startColumnNumber, rowNumber, rateCard.Item1.Value, currentActiveWorksheet);
                    //rowNumber += rateCard.meters.Count;
                }
                else
                {
                    var rateCard = await BillingUtils.GetRateCardStandardAsync(token, subscriptionId, offerDurableId, currency, locale, regionInfo);
                    if (rateCard?.Item1 == null)
                    {
                        MessageBox.Show(
                            $"ERROR: Failed to get the rate card. Verify the correct parameters were provided for Subscription Id, Offer Id, currency, locale, and region info and try again.",
                            $"Get Rate Card ({usageApi})", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Add a worksheet for raw payload.
                    if (this.IncludeRawPayloadCheckBox.Checked && !string.IsNullOrWhiteSpace(rateCard.Item2))
                    {
                        ShowRawPayload("ratecard-std-", FormatJson(rateCard.Item2));
                    }

                    // Add a fresh worksheet and write the results.
                    Excel.Worksheet currentActiveWorksheet =
                        Globals.ThisAddIn.Application.Worksheets.Add(Globals.ThisAddIn.Application.ActiveSheet);
                    currentActiveWorksheet.SetWorksheetName(usageApi, BillingApiType.RateCard);
                    var rowNumber = this.PrintRateCardHeader(startColumnNumber, startHeaderRowNumber, rateCard.Item1.Value,
                        currentActiveWorksheet, usageApi);
                    this.PrintRateCardReport(startColumnNumber, rowNumber, rateCard.Item1.Value, currentActiveWorksheet);
                    //rowNumber += rateCard.Item1.Meters.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to get rate card: {ex.Message}\r\n\r\n\r\n{ex.StackTrace}\r\n",
                    $"Get Rate Card ({usageApi})", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private async Task GetPriceSheetAsync()
        {
            if (!this.ValidateUsageReportInput(UsageApi.EnterpriseAgreement))
            {
                return;
            }

            try
            {
                Globals.ThisAddIn.Application.StatusBar = $"Getting price sheet (EA)...";

                var enrollmentNumber = this.EnrollmentNumberComboBox.Text.Trim();
                var billingPeriod = this.PriceSheetBillingPeriodComboBox.Text.Trim();
                var apiKey = this.EaApiKeyEditBox.Text.Trim();

                // Write the report line items:
                int startColumnNumber = 1; // A
                int startHeaderRowNumber = 1;

                var priceSheet = await BillingUtils.GetPriceSheetAsync(apiKey, enrollmentNumber, billingPeriod);
                if (priceSheet?.Item1 == null)
                {
                    MessageBox.Show(
                        $"ERROR: Failed to get the price sheet. Verify the correct parameters were provided for enrollment number and billing period and try again.",
                        $"Get Rate Card (EA)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Add a worksheet for raw payload.
                if (this.IncludeRawPayloadCheckBox.Checked && !string.IsNullOrWhiteSpace(priceSheet.Item2))
                {
                    ShowRawPayload("pricesheet-ea-", FormatJson(priceSheet.Item2));
                }

                // Add a fresh worksheet and write the results.
                Excel.Worksheet currentActiveWorksheet =
                    Globals.ThisAddIn.Application.Worksheets.Add(Globals.ThisAddIn.Application.ActiveSheet);
                currentActiveWorksheet.SetWorksheetName(UsageApi.EnterpriseAgreement, BillingApiType.RateCard);
                var rowNumber = this.PrintPriceSheetHeader(startColumnNumber, startHeaderRowNumber, priceSheet.Item1.Value,
                    currentActiveWorksheet, UsageApi.EnterpriseAgreement);
                this.PrintPriceSheetReport(startColumnNumber, rowNumber, priceSheet.Item1.Value, currentActiveWorksheet);
                //rowNumber += priceSheet.Item1.Meters.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to get price sheet: {ex.Message}\r\n\r\n\r\n{ex.StackTrace}\r\n",
                    $"Get Rate Card (EA)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private static string FormatJson(string jsonPayload)
        {
            dynamic parsedJson = JsonConvert.DeserializeObject(jsonPayload);
            return JsonConvert.SerializeObject(parsedJson, Formatting.Indented);
        }

        private static void ShowRawPayload(string prefix, string payload)
        {
            string fileName = $"{prefix}{new Random((int)DateTime.Now.Ticks).Next(10000000, 99999999)}.json";
            string pathToPayload = Path.Combine(Environment.GetEnvironmentVariable("TEMP"), fileName);
            File.WriteAllText(pathToPayload, payload);
            Process.Start("notepad.exe", pathToPayload);
        }

        private int PrintUsageAggregatesHeader(int startColumnNumber, int headerRowNumber, Excel.Worksheet currentActiveWorksheet, UsageApi usageApi)
        {
            Globals.ThisAddIn.Application.StatusBar = "Displaying usage report header...";

            var tenantId = this.TenantIdComboBox.Text.Trim();
            var subscriptionId = this.SubscriptionIdComboBox.Text.Trim();
            var reportStartDate = this.StartDateEditBox.Text.Trim();
            var reportEndDate = this.EndDateEditBox.Text.Trim();
            var aggregationGranularity = this.AggregationGranularityDropDown.SelectedItem.Tag;
            var enrollmentNumber = this.EnrollmentNumberComboBox.Text.Trim();
            var billingPeriod = this.PriceSheetBillingPeriodComboBox.Text.Trim();

            // Write the report header:
            int rowNumber = headerRowNumber;
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Subscription Id:", $"{subscriptionId}" }, currentActiveWorksheet);
            if (usageApi == UsageApi.Standard || usageApi == UsageApi.CloudSolutionProvider)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Tenant Id:", $"{tenantId}" }, currentActiveWorksheet);
            }
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Report Start Date (UTC):", $"{reportStartDate}" }, currentActiveWorksheet);
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Report End Date (UTC):", $"{reportEndDate}" }, currentActiveWorksheet);
            if (usageApi == UsageApi.Standard || usageApi == UsageApi.CloudSolutionProvider)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Aggregation Granularity:", $"{aggregationGranularity}" }, currentActiveWorksheet);
            }
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Report generated (UTC):", $"{DateTime.UtcNow}" }, currentActiveWorksheet);
            if (usageApi == UsageApi.EnterpriseAgreement)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Enrollment Number:", $"{enrollmentNumber}" }, currentActiveWorksheet);
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Billing Period:", $"{billingPeriod}" }, currentActiveWorksheet);
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
                if (rateCard.OfferTerms[0].ExcludedMeterIds != null)
                {
                    ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Excluded Meter Ids:", $"{string.Join(";", rateCard.OfferTerms[0].ExcludedMeterIds)}" }, currentActiveWorksheet);
                }
                else
                {
                    ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Excluded Meter Ids:", string.Empty }, currentActiveWorksheet);
                }

                var tieredDiscount = rateCard.OfferTerms[0].TieredDiscount;
                if (tieredDiscount != null)
                {
                    ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Terms Duration:", $"{string.Join(";", tieredDiscount.Select(x => x.Key + "=" + x.Value))}" }, currentActiveWorksheet);
                }
                else
                {
                    ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Terms Duration:", string.Empty }, currentActiveWorksheet);
                }
            }

            ExcelUtils.WriteUsageLineItemHeader(startColumnNumber, rowNumber, this.GetRateCardHeaderCaptions(usageApi), currentActiveWorksheet);

            // Format the data types of datatime and numeric columns.
            rowNumber++; // Starting row number for the table.
            currentActiveWorksheet.get_Range($"J{rowNumber}").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // EffectiveDate

            return ++rowNumber; // return the first writable row number.
        }

        private int PrintCspRateCardHeader(int startColumnNumber, int headerRowNumber, CspRateCard rateCard, Excel.Worksheet currentActiveWorksheet, UsageApi usageApi)
        {
            Globals.ThisAddIn.Application.StatusBar = "Displaying rate card header (CSP)...";

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
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Is Tax Included:", $"{rateCard.isTaxIncluded}" }, currentActiveWorksheet);
            if (rateCard.tags != null)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "RateCard Tags:", $"{string.Join(";", rateCard.tags)}" }, currentActiveWorksheet);
            }
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "RateCard generated (UTC):", $"{DateTime.UtcNow}" }, currentActiveWorksheet);
            if (rateCard.offerTerms.Count > 0)
            {
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Terms Name:", $"{rateCard.offerTerms[0].name}" }, currentActiveWorksheet);
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Offer Terms Duration:", $"{rateCard.offerTerms[0].effectiveDate}" }, currentActiveWorksheet);
                ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Discount:", $"{rateCard.offerTerms[0].discount}" }, currentActiveWorksheet);
                if (rateCard.offerTerms[0].excludedMeterIds != null)
                {
                    ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Excluded Meter Ids:", $"{string.Join(";", rateCard.offerTerms[0].excludedMeterIds)}" }, currentActiveWorksheet);
                }
                else
                {
                    ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Excluded Meter Ids:", string.Empty }, currentActiveWorksheet);
                }
            }

            ExcelUtils.WriteUsageLineItemHeader(startColumnNumber, rowNumber, this.GetRateCardHeaderCaptions(usageApi), currentActiveWorksheet);

            // Format the data types of datatime and numeric columns.
            rowNumber++; // Starting row number for the table.
            currentActiveWorksheet.get_Range($"J{rowNumber}").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // EffectiveDate

            return ++rowNumber; // return the first writable row number.
        }

        private int PrintPriceSheetHeader(int startColumnNumber, int headerRowNumber, PriceSheet priceSheet, Excel.Worksheet currentActiveWorksheet, UsageApi usageApi)
        {
            Globals.ThisAddIn.Application.StatusBar = "Displaying price sheet header...";

            var billingPeriod = this.PriceSheetBillingPeriodComboBox.Text.Trim();

            // Write the report header:
            int rowNumber = headerRowNumber;
            ExcelUtils.WriteHeaderRow($"A{rowNumber}", $"B{rowNumber++}", new[] { "Billing Period:", $"{billingPeriod}" }, currentActiveWorksheet);
            ExcelUtils.WriteUsageLineItemHeader(startColumnNumber, rowNumber++, this.GetRateCardHeaderCaptions(usageApi), currentActiveWorksheet);

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
                    return this.HeaderCaptionsPriceSheetMeter;
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

        private void PrintCspRateCardReport(int startColumnNumber, int rowNumber, CspRateCard rateCard, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying CSP rate card. Please wait...";

            if (rateCard.meters != null)
            {
                foreach (var meter in rateCard.meters)
                {
                    ExcelUtils.WriteCspRateCardMeterLineItem(startColumnNumber, rowNumber, meter,
                        this.HeaderCaptionsRateCardMeter.Length, currentActiveWorksheet);
                    rowNumber++;
                }
            }
        }

        private void PrintPriceSheetReport(int startColumnNumber, int rowNumber, PriceSheet priceSheet, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying EA price sheet. Please wait...";

            if (priceSheet.PriceSheetMeters != null)
            {
                foreach (var meter in priceSheet.PriceSheetMeters)
                {
                    ExcelUtils.WritePriceSheetMeterLineItem(startColumnNumber, rowNumber, meter,
                        this.HeaderCaptionsPriceSheetMeter.Length, currentActiveWorksheet);
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
                EaApiKey = this.EaApiKeyEditBox.Text.Trim(),
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
                    this.EaApiKeyEditBox.Text = persistedData.EaApiKey;
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
            this.EaApiKeyEditBox.Enabled = isEa;
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

            if (string.IsNullOrWhiteSpace(this.StartDateEditBox.Text) && this.StartDateEditBox.Enabled && (usageApi == UsageApi.CloudSolutionProvider || usageApi == UsageApi.Standard))
            {
                MessageBox.Show($"ERROR: Report Start Date (yyyy-mm-dd) must be specified.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.EndDateEditBox.Text) && this.EndDateEditBox.Enabled && (usageApi == UsageApi.CloudSolutionProvider || usageApi == UsageApi.Standard))
            {
                MessageBox.Show($"ERROR: Report End Date (yyyy-mm-dd) must be specified.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.EnrollmentNumberComboBox.Text) && this.EnrollmentNumberComboBox.Enabled)
            {
                MessageBox.Show($"ERROR: Enrollment Number must be specified for an EA Usage Report.", "Azure Excel Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.EaApiKeyEditBox.Text) && this.EaApiKeyEditBox.Enabled)
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
