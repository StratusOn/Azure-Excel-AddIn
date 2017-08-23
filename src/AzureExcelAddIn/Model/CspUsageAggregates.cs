using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ExcelAddIn1
{
    public class CspUsageAggregates
    {
        public int totalCount { get; set; }
        public List<Item> items { get; set; }
        public Links links { get; set; }
        public Attributes attributes { get; set; }
    }

    public class Links
    {
        public Self self { get; set; }
    }

    public class Self
    {
        public string uri { get; set; }
        public string method { get; set; }
        public object[] headers { get; set; }
    }

    public class Attributes
    {
        public string objectType { get; set; }
    }

    public class Item
    {
        public DateTime usageStartTime { get; set; }
        public DateTime usageEndTime { get; set; }
        public Resource resource { get; set; }
        public double quantity { get; set; }
        public string unit { get; set; }
        public InfoFields infoFields { get; set; }
        [JsonProperty("instanceData")]
        public string instanceDataRaw { get; set; }
        public InstanceData InstanceData => JsonConvert.DeserializeObject<InstanceData>(instanceDataRaw.Replace("\\\"", ""));
        public Attributes attributes { get; set; }
    }

    public class Resource
    {
        public string id { get; set; }
        public string name { get; set; }
        public string category { get; set; }
        public string subcategory { get; set; }
        public string region { get; set; }
    }
}
