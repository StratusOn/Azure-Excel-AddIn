using System;

namespace ExcelAddIn1
{
    public class EaUsageAggregates
    {
        public string id { get; set; }
        public Datum[] data { get; set; }
        public string nextLink { get; set; }
    }

    public class Datum
    {
        public int accountId { get; set; }
        public int productId { get; set; }
        public int resourceLocationId { get; set; }
        public int consumedServiceId { get; set; }
        public int departmentId { get; set; }
        public string accountOwnerEmail { get; set; }
        public string accountName { get; set; }
        public string serviceAdministratorId { get; set; }
        public int subscriptionId { get; set; }
        public string subscriptionGuid { get; set; }
        public string subscriptionName { get; set; }
        public DateTime date { get; set; }
        public string product { get; set; }
        public string meterId { get; set; }
        public string meterCategory { get; set; }
        public string meterSubCategory { get; set; }
        public string meterRegion { get; set; }
        public string meterName { get; set; }
        public int consumedQuantity { get; set; }
        public int resourceRate { get; set; }
        public int Cost { get; set; }
        public string resourceLocation { get; set; }
        public string consumedService { get; set; }
        public string instanceId { get; set; }
        public string serviceInfo1 { get; set; }
        public string serviceInfo2 { get; set; }
        public string additionalInfo { get; set; }
        public string tags { get; set; }
        public string storeServiceIdentifier { get; set; }
        public string departmentName { get; set; }
        public string costCenter { get; set; }
        public string unitOfMeasure { get; set; }
        public string resourceGroup { get; set; }
    }

}
