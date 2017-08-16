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
        public Offerterm[] OfferTerms { get; set; }
        public Meter[] Meters { get; set; }
        public string Currency { get; set; }
        public string Locale { get; set; }
        public bool IsTaxIncluded { get; set; }
        public string MeterRegion { get; set; }
        public object[] Tags { get; set; }
    }

    public class Offerterm
    {
        public string Name { get; set; }
        public float Credit { get; set; }
        public JObject TieredDiscount { get; set; }
        public object[] ExcludedMeterIds { get; set; }
        public DateTime EffectiveDate { get; set; }
    }

  
    public class Meter
    {
        public string MeterId { get; set; }
        public string MeterName { get; set; }
        public string MeterCategory { get; set; }
        public string MeterSubCategory { get; set; }
        public string Unit { get; set; }
        public object[] MeterTags { get; set; }
        public string MeterRegion { get; set; }
        public JObject MeterRates { get; set; }
        public DateTime EffectiveDate { get; set; }
        public float IncludedQuantity { get; set; }
    }
}
