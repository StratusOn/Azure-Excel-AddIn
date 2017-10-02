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
        public LinkItem self { get; set; }
        public LinkItem next { get; set; }
    }

    public class LinkItem
    {
        public string uri { get; set; }
        public string method { get; set; }
        public List<HeaderItem> headers { get; set; }
    }

    public class HeaderItem
    {
        public string key { get; set; }
        public string value { get; set; }
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
        public CspInstanceData InstanceData { get; set; }
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
