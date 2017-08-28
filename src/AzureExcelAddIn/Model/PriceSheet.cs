namespace ExcelAddIn1
{
    public class PriceSheet
    {
        public PriceSheetMeter[] PriceSheetMeters { get; set; }
    }

    public class PriceSheetMeter
    {
        public string id { get; set; }
        public string billingPeriodId { get; set; }
        public string meterId { get; set; }
        public string meterName { get; set; }
        public string meterRegion { get; set; }
        public string unitOfMeasure { get; set; }
        public double includedQuantity { get; set; }
        public string partNumber { get; set; }
        public double unitPrice { get; set; }
        public string currencyCode { get; set; }
    }

}
