using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class AzureRibbon
    {
        private const int MaxContinuationLinks = 500;

        private readonly string[] HeaderCaptions = {
            "Id", "Name", "Type", "subscription Id", "Usage Start Time (UTC)", "Usage End Time (UTC)", "Meter Id", "Meter Name",
            "Meter Category", "Meter Sub-Category", "Quantity", "Unit", "Tags", "Info Fields (legacy format)", "Instance Data (new format)"
        };

        private void AzureRibbonTab_Load(object sender, RibbonUIEventArgs e)
        {
            var today = DateTime.Today;
            var yesterday = today.AddDays(-1);
            this.StartDateEditBox.Text = $"{yesterday.Year}-{yesterday.Month:0#}-{yesterday.Day:0#}";
            this.EndDateEditBox.Text = $"{today.Year}-{today.Month:0#}-{today.Day:0#}";
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
                string token = AuthUtils.GetAuthorizationHeaderAsync(tenantId, true);

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
            if (!this.ValidateUsageReportInput())
            {
                return;
            }

            var tenantId = this.TenantIdEditBox.Text;

            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Authenticating...";

                string token = AuthUtils.GetAuthorizationHeaderAsync(tenantId, this.ForceReAuthCheckBox.Checked);

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show(
                        $"ERROR: Failed to acquire a token. Verify you entered the right credentials, the correct Tenant Id and Subscription Id, and make sure 'Force Re-Authentication' is checked and try again.",
                        "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Globals.ThisAddIn.Application.StatusBar = "Getting usage report...";

                var subscriptionId = this.SubscriptionIdEditBox.Text;
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
                UsageAggregates usageAggregates = await BillingUtils.GetUsageAggregates(token, subscriptionId, reportStartDate, reportEndDate, aggregationGranularity, showDetails);

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
                        this.PrintUsageAggregatesHeader(startColumnNumber, headerRowNumber, currentActiveWorksheet);
                    }

                    this.PrintUsageAggregatesReport(startColumnNumber, rowNumber, usageAggregates, currentContinuationCount, currentActiveWorksheet);
                    rowNumber += usageAggregates.value.Length;

                    // A maximum of 1000 records are returned by the API. If more than 1000 records will be returned, a continuation link is provided to get the next chunk and so on.
                    string continuationLink = usageAggregates.nextLink;
                    if (string.IsNullOrWhiteSpace(continuationLink))
                    {
                        break;
                    }

                    usageAggregates = await BillingUtils.GetUsageAggregates(token, continuationLink);
                    currentContinuationCount++;
                } while (currentContinuationCount < MaxContinuationLinks);

                currentActiveWorksheet.get_Range($"M11:M{rowNumber - 1}").Interior.ColorIndex = 19; // #FFFFCC // Tags
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERROR: Failed to get usage report: {ex.Message}\r\n\r\n{ex.StackTrace}\r\n",
                    "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = "Ready";
            }
        }

        private bool ValidateUsageReportInput()
        {
            if (string.IsNullOrWhiteSpace(this.TenantIdEditBox.Text))
            {
                MessageBox.Show($"ERROR: Tenant Id must be specified.", "Get Usage Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(this.SubscriptionIdEditBox.Text))
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

            return true;
        }

        private void PrintUsageAggregatesHeader(int startColumnNumber, int headerRowNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = "Displaying usage report header...";

            var tenantId = this.TenantIdEditBox.Text;
            var subscriptionId = this.SubscriptionIdEditBox.Text;
            var reportStartDate = this.StartDateEditBox.Text;
            var reportEndDate = this.EndDateEditBox.Text;
            var aggregationGranularity = this.AggregationGranularityDropDown.SelectedItem.Tag;


            // Write the report header:
            this.WriteHeaderRow("A1", "B1", new[] { "Subscription Id:", $"{subscriptionId}" }, currentActiveWorksheet);
            this.WriteHeaderRow("A2", "B2", new[] { "Tenant Id:", $"{tenantId}" }, currentActiveWorksheet);
            this.WriteHeaderRow("A3", "B3", new[] { "Report Start Date (UTC):", $"{reportStartDate}" }, currentActiveWorksheet);
            this.WriteHeaderRow("A4", "B4", new[] { "Report End Date (UTC):", $"{reportEndDate}" }, currentActiveWorksheet);
            this.WriteHeaderRow("A5", "B5", new[] { "Aggregation Granularity:", $"{aggregationGranularity}" }, currentActiveWorksheet);
            this.WriteHeaderRow("A6", "B6", new[] { "Report generated (UTC):", $"{DateTime.UtcNow}" }, currentActiveWorksheet);


            this.WriteUsageLineItemHeader(startColumnNumber, headerRowNumber, this.HeaderCaptions, currentActiveWorksheet);

            // Format the data types of datatime and numeric columns.
            currentActiveWorksheet.get_Range("E11").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // UsageStartTime
            currentActiveWorksheet.get_Range("F11").EntireColumn.NumberFormat = "yyyy-mm-dd HH:mm:ss"; // UsageEndTime
            //currentActiveWorksheet.get_Range("K11").EntireColumn.NumberFormat = "#####.###################"; // Quantity
        }

        private void PrintUsageAggregatesReport(int startColumnNumber, int rowNumber, UsageAggregates usageAggregates, int chunkNumber, Excel.Worksheet currentActiveWorksheet)
        {
            Globals.ThisAddIn.Application.StatusBar = $"Displaying usage report chunk {chunkNumber}...";

            foreach (var usageAggregate in usageAggregates.value)
            {
                this.WriteUsageLineItem(startColumnNumber, rowNumber, usageAggregate, this.HeaderCaptions.Length, currentActiveWorksheet);
                rowNumber++;
            }
        }

        private void WriteHeaderRow(string nameCell, string valueCell, string[] content, Excel.Worksheet activeWorksheet)
        {
            Excel.Range currentRow = activeWorksheet.get_Range(nameCell);
            currentRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newCurrentRow = activeWorksheet.get_Range(nameCell, valueCell);
            newCurrentRow.Value2 = content;
            newCurrentRow.Interior.ColorIndex = 36; // #FFFF99
            newCurrentRow.ColumnWidth = 35;
        }

        private void WriteUsageLineItemHeader(int startColumnNumber, int rowNumber, string[] headerCaptions, Excel.Worksheet activeWorksheet)
        {
            Excel.Range c1 = (Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Excel.Range c2 = (Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + headerCaptions.Length - 1];
            Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);
            currentRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newCurrentRow = activeWorksheet.get_Range(c1, c2);
            newCurrentRow.Value2 = headerCaptions;
            newCurrentRow.Interior.ColorIndex = 15; // #C0C0C0
            newCurrentRow.ColumnWidth = 35;
        }

        private void WriteUsageLineItem(int startColumnNumber, int rowNumber, Value lineItem, int numberOfColumns, Excel.Worksheet activeWorksheet)
        {
            Excel.Range c1 = (Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Excel.Range c2 = (Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);
            currentRow.Value2 = this.GetLineItemFields(lineItem);
        }

        private object[] GetLineItemFields(Value lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.id);
            fields.Add(lineItem.name);
            fields.Add(lineItem.type);
            fields.Add(lineItem.properties.subscriptionId);
            fields.Add(lineItem.properties.usageStartTime);
            fields.Add(lineItem.properties.usageEndTime);
            fields.Add(lineItem.properties.meterId);
            fields.Add(lineItem.properties.meterName);
            fields.Add(lineItem.properties.meterCategory);
            fields.Add(lineItem.properties.meterSubCategory);
            fields.Add(lineItem.properties.quantity);
            fields.Add(lineItem.properties.unit);
            fields.Add(JsonUtils.ExtractTagsFromInstanceData(lineItem.properties.instanceData));
            fields.Add(JsonUtils.ExtractInfoFields(lineItem.properties.infoFields));
            fields.Add(lineItem.properties.instanceData);

            return fields.ToArray();
        }
    }
}
