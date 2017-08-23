using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ExcelAddIn1
{
    // Reference: https://github.com/Azure-Samples/billing-dotnet-usage-api.
    public class UsageAggregates
    {
        public List<Value> value { get; set; }
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
        public string meterRegion { get; set; }
        public string meterCategory { get; set; }
        public string meterSubCategory { get; set; }
        public string unit { get; set; }
        public string meterId { get; set; }
        public InfoFields infoFields { get; set; }
        public double quantity { get; set; }
        [JsonProperty("instanceData")]
        public string instanceDataRaw { get; set; }
        public InstanceData InstanceData
        {
            get
            {
                if (instanceDataRaw != null)
                {
                    return JsonConvert.DeserializeObject<InstanceData>(instanceDataRaw.Replace("\\\"", ""));
                }
                return null;
            }
        }
    }
}
