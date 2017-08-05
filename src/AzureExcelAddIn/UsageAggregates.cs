using System;
using Newtonsoft.Json.Linq;

namespace ExcelAddIn1
{
    public class UsageAggregates
    {
        public Value[] value { get; set; }
        public string nextLink { get; set; }
    }

    public class Value
    {
        public string id { get; set; }
        public string name { get; set; }
        public string type { get; set; }
        public PropertiesEx properties { get; set; }
    }

    public class PropertiesEx
    {
        public string subscriptionId { get; set; }
        public DateTime usageStartTime { get; set; }
        public DateTime usageEndTime { get; set; }
        public string meterName { get; set; }
        public string meterCategory { get; set; }
        public string meterSubCategory { get; set; }
        public string unit { get; set; }
        public string instanceData { get; set; }
        public string meterId { get; set; }
        public JObject infoFields { get; set; }
        public float quantity { get; set; }
    }
}
