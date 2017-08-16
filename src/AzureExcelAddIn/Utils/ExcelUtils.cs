namespace ExcelAddIn1
{
    internal static class ExcelUtils
    {
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

        public static void WriteRateCardLineItem(int startColumnNumber, int rowNumber, Meter meterItem, int numberOfColumns, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.Cells[rowNumber, startColumnNumber + numberOfColumns - 1];
            Microsoft.Office.Interop.Excel.Range currentRow = activeWorksheet.get_Range(c1, c2);

            currentRow.Value2 = BillingUtils.GetLineItemFields(meterItem);
        }
    }
}
