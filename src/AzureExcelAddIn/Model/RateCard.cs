using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelAddIn1
{
    public class RateCard
    {
        public List<Offerterm> OfferTerms { get; set; }
        public List<Meter> Meters { get; set; }
        public string Currency { get; set; }
        public string Locale { get; set; }
        public bool IsTaxIncluded { get; set; }
        public string MeterRegion { get; set; }
        public List<string> Tags { get; set; }
    }

    public class Offerterm
    {
        public string Name { get; set; }
        public double Credit { get; set; }
        public IDictionary<string, double> TieredDiscount { get; set; }
        public List<string> ExcludedMeterIds { get; set; }
        public DateTime EffectiveDate { get; set; }
    }


    public class Meter
    {
        public string MeterId { get; set; }
        public string MeterName { get; set; }
        public string MeterCategory { get; set; }
        public string MeterSubCategory { get; set; }
        public string Unit { get; set; }
        public List<string> MeterTags { get; set; }
        public string MeterRegion { get; set; }
        public IDictionary<string, double> MeterRates { get; set; }
        public DateTime EffectiveDate { get; set; }
        public double IncludedQuantity { get; set; }
        public string MeterStatus { get; set; }
    }
}
