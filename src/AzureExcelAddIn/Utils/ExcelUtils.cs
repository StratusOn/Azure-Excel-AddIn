using System;
using System.Globalization;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public enum BillingApiType
    {
        Usage = 0,
        RateCard = 1
    }

    internal static class ExcelUtils
    {
        private const string StandardUsageReportWorksheetNameTemplate = "Azure Usage{0}";
        private const string CspUsageReportWorksheetNameTemplate = "Azure CSP Usage{0}";
        private const string EaUsageReportWorksheetNameTemplate = "Azure EA Usage{0}";
        private const string StandardRateCardWorksheetNameTemplate = "RateCard{0}";
        private const string CspRateCardWorksheetNameTemplate = "CSP RateCard{0}";
        private const string EaRateCardWorksheetNameTemplate = "EA RateCard{0}";
        private const int MaxNameRetries = 100;

        public static void WriteHeaderRow(string nameCell, string valueCell, string[] content, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(nameCell);
            currentRow.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            Microsoft.Office.Interop.Excel.Range newCurrentRow = activeWorksheet.get_Range(nameCell, valueCell);
            newCurrentRow.Value2 = content;
            newCurrentRow.Interior.ColorIndex = 36; // #FFFF99
            newCurrentRow.ColumnWidth = 35;
        }

        public static void WriteUsageLineItemHeader(int startColumnNumber, int rowNumber, string[] headerCaptions, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + headerCaptions.Length - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);
            currentRow.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            Microsoft.Office.Interop.Excel.Range newCurrentRow = activeWorksheet.get_Range(c1, c2);
            newCurrentRow.Value2 = headerCaptions;
            newCurrentRow.Interior.ColorIndex = 15; // #C0C0C0
            newCurrentRow.ColumnWidth = 35;
        }

        public static void WriteUsageLineItem(int startColumnNumber, int rowNumber, Value lineItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);

            currentRow.Value2 = BillingUtils.GetLineItemFields(lineItem);
        }

        public static void WriteUsageLineItemCsp(int startColumnNumber, int rowNumber, Item lineItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);
            currentRow.Value2 = BillingUtils.GetLineItemFieldsCsp(lineItem);
        }

        public static void WriteUsageLineItemEa(int startColumnNumber, int rowNumber, Datum lineItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);
            currentRow.Value2 = BillingUtils.GetLineItemFieldsEa(lineItem);
        }

        public static void WriteRateCardOfferTermLineItem(int startColumnNumber, int rowNumber, Offerterm offerTermItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);

            currentRow.Value2 = BillingUtils.GetRateCardOfferTermLineItemFields(offerTermItem);
        }

        public static void WriteRateCardMeterLineItem(int startColumnNumber, int rowNumber, Meter meterItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);

            currentRow.Value2 = BillingUtils.GetRateCardMeterLineItemFields(meterItem);
        }

        public static void WriteCspRateCardMeterLineItem(int startColumnNumber, int rowNumber, CspMeter meterItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);

            currentRow.Value2 = BillingUtils.GetCspRateCardMeterLineItemFields(meterItem);
        }

        public static void WritePriceSheetMeterLineItem(int startColumnNumber, int rowNumber, PriceSheetMeter meterItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);

            currentRow.Value2 = BillingUtils.GetPriceSheetMeterLineItemFields(meterItem);
        }

        public static void SetWorksheetName(this Worksheet worksheet, UsageApi usageApi, BillingApiType billingApiType)
        {
            int counter = 1;
            do
            {
                string worksheetName = null;
                switch (usageApi)
                {
                    case UsageApi.CloudSolutionProvider:
                        worksheetName = string.Format(CultureInfo.CurrentUICulture,
                            billingApiType == BillingApiType.RateCard
                                ? CspRateCardWorksheetNameTemplate
                                : CspUsageReportWorksheetNameTemplate, counter);
                        break;
                    case UsageApi.EnterpriseAgreement:
                        worksheetName = string.Format(CultureInfo.CurrentUICulture,
                            billingApiType == BillingApiType.RateCard
                                ? EaRateCardWorksheetNameTemplate
                                : EaUsageReportWorksheetNameTemplate, counter);
                        break;
                    default:
                        worksheetName = string.Format(CultureInfo.CurrentUICulture,
                            billingApiType == BillingApiType.RateCard
                                ? StandardRateCardWorksheetNameTemplate
                                : StandardUsageReportWorksheetNameTemplate, counter);
                        break;
                }

                if (!Globals.ThisAddIn.Application.Worksheets.Contains(worksheetName))
                {
                    worksheet.Name = worksheetName;
                    return;
                }

                counter++;
            } while (counter < MaxNameRetries);
        }

        public static bool Contains(this Sheets sheets, string name)
        {
            foreach (Worksheet sheet in sheets)
            {
                if (string.Compare(sheet.Name, name, StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
