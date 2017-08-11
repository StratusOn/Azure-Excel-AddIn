using System;
using Newtonsoft.Json.Linq;

namespace ExcelAddIn1
{
    public class CspUsageAggregates
    {
        public int totalCount { get; set; }
        public Item[] items { get; set; }
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
        public float quantity { get; set; }
        public string unit { get; set; }
        public Infofields infoFields { get; set; }
        public Instancedata instanceData { get; set; }
        public Attributes1 attributes { get; set; }
    }

    public class Resource
    {
        public string id { get; set; }
        public string name { get; set; }
        public string category { get; set; }
        public string subcategory { get; set; }
        public string region { get; set; }
    }

    public class Infofields : JObject
    {
    }

    public class Instancedata
    {
        public string resourceUri { get; set; }
        public string location { get; set; }
        public string partNumber { get; set; }
        public string orderNumber { get; set; }
    }

    public class Attributes1
    {
        public string objectType { get; set; }
    }

}
