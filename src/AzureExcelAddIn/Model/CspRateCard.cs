using System;
using System.Collections.Generic;

namespace ExcelAddIn1
{
    public class CspRateCard
    {
        public List<CspOfferterm> offerTerms { get; set; }
        public List<CspMeter> meters { get; set; }
        public string currency { get; set; }
        public string locale { get; set; }
        public bool isTaxIncluded { get; set; }
        public string meterRegion { get; set; }
        public List<string> tags { get; set; }
    }

    public class CspOfferterm
    {
        public string name { get; set; }
        public double discount { get; set; }
        public List<string> excludedMeterIds { get; set; }
        public DateTime effectiveDate { get; set; }
    }


    public class CspMeter
    {
        public string id { get; set; }
        public string name { get; set; }
        public string category { get; set; }
        public string subcategory { get; set; }
        public string unit { get; set; }
        public List<string> tags { get; set; }
        public string region { get; set; }
        public IDictionary<string, double> rates { get; set; }
        public DateTime effectiveDate { get; set; }
        public double includedQuantity { get; set; }
        public string status { get; set; }
    }
}
