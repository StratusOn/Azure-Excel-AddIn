using System;
using System.Collections.Generic;

namespace ExcelAddIn1
{
    public class EaUsageAggregates
    {
        public string id { get; set; }
        public List<Datum> data { get; set; }
        public string nextLink { get; set; }
    }

    public class Datum
    {
        public long accountId { get; set; }
        public long productId { get; set; }
        public long resourceLocationId { get; set; }
        public long consumedServiceId { get; set; }
        public long departmentId { get; set; }
        public string accountOwnerEmail { get; set; }
        public string accountName { get; set; }
        public string serviceAdministratorId { get; set; }
        public long subscriptionId { get; set; }
        public string subscriptionGuid { get; set; }
        public string subscriptionName { get; set; }
        public DateTime date { get; set; }
        public string product { get; set; }
        public string meterId { get; set; }
        public string meterCategory { get; set; }
        public string meterSubCategory { get; set; }
        public string meterRegion { get; set; }
        public string meterName { get; set; }
        public double consumedQuantity { get; set; }
        public double resourceRate { get; set; }
        public double Cost { get; set; }
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
