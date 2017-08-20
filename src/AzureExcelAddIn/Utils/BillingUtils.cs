using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelAddIn1
{
    internal class BillingUtils
    {
        public static async Task<RateCard> GetRateCardStandardAsync(string authorizationToken, string subscriptionId, string offerDurableId, string currency, string locale, string regionInfo, string apiVersion = "2015-06-01-preview")
        {
            string usageAggregatesUrl =
            $"https://management.azure.com/subscriptions/{subscriptionId}/providers/Microsoft.Commerce/RateCard?api-version={apiVersion}&$filter=OfferDurableId eq '{offerDurableId}' and Currency eq '{currency}' and Locale eq '{locale}' and RegionInfo eq '{regionInfo}'";
            string content = await GetRestCallResultsAsync(authorizationToken, usageAggregatesUrl);
            return JsonConvert.DeserializeObject<RateCard>(content);
        }

        public static async Task<UsageAggregates> GetUsageAggregatesStandardAsync(string authorizationToken, string subscriptionId, string reportStartDate, string reportEndDate, string aggregationGranularity, string showDetails)
        {
            string usageAggregatesUrl =
                        $"https://management.azure.com/subscriptions/{subscriptionId}/providers/Microsoft.Commerce/UsageAggregates?api-version=2015-06-01-preview&reportedStartTime={reportStartDate}&reportedEndTime={reportEndDate}&aggregationGranularity={aggregationGranularity}&showDetails={showDetails}";
            string content = await GetRestCallResultsAsync(authorizationToken, usageAggregatesUrl);
            return JsonConvert.DeserializeObject<UsageAggregates>(content);
        }

        public static async Task<CspUsageAggregates> GetUsageAggregatesCspAsync(string authorizationToken, string subscriptionId, string customerTenantId, string reportStartDate, string reportEndDate, string aggregationGranularity, string showDetails, int chunkSize)
        {
            string usageAggregatesUrl =
                $"https://api.partnercenter.microsoft.com/v1/customers/{customerTenantId}/subscriptions/{subscriptionId}/utilizations/azure?start_time={reportStartDate}&end_time={reportEndDate}&size={chunkSize}";
            string content = await GetRestCallResultsAsync(authorizationToken, usageAggregatesUrl);
            return JsonConvert.DeserializeObject<CspUsageAggregates>(content);
        }

        public static async Task<EaUsageAggregates> GetUsageAggregatesEaAsync(string authorizationToken, string enrollmentNumber, string reportStartDate, string reportEndDate)
        {
            string usageAggregatesUrl = $"https://consumption.azure.com/v2/enrollments/{enrollmentNumber}/usagedetailsbycustomdate?startTime={reportStartDate}&endTime={reportEndDate}";
            string content = await GetRestCallResultsAsync(authorizationToken, usageAggregatesUrl);
            return JsonConvert.DeserializeObject<EaUsageAggregates>(content);
        }

        public static async Task<string> GetRestCallResultsAsync(string authorizationToken, string url)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            HttpResponseMessage response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }

            string statusCodeName = response.StatusCode.ToString();
            int statusCodeValue = (int)response.StatusCode;
            string content = await response.Content.ReadAsStringAsync();
            throw new Exception($"Status Code: {statusCodeName} ({statusCodeValue}). \r\nBody:\r\n{content}");
        }

        public static string CreateStartTime(string startDate)
        {
            if (startDate.Contains("T"))
            {
                return startDate;
            }

            return $"{startDate}T00%3a00%3a00%2b00%3a00";
        }

        public static string CreateEndTime(string endDate)
        {
            if (endDate.Contains("T"))
            {
                return endDate;
            }

            return $"{endDate}T00%3a00%3a00%2b00%3a00";
        }

        public static object[] GetLineItemFields(Value lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.properties.usageStartTime);
            fields.Add(lineItem.properties.usageEndTime);
            fields.Add(lineItem.id);
            fields.Add(lineItem.name);
            fields.Add(lineItem.type);
            fields.Add(lineItem.properties.subscriptionId);
            fields.Add(lineItem.properties.meterId);
            fields.Add(lineItem.properties.meterName);
            fields.Add(lineItem.properties.meterCategory);
            fields.Add(lineItem.properties.meterSubCategory);
            fields.Add(lineItem.properties.quantity);
            fields.Add(lineItem.properties.unit);
            fields.Add(JsonUtils.ExtractTagsFromInstanceData(lineItem.properties.instanceData));
            fields.Add(JsonUtils.ExtractInfoFields(lineItem.properties.infoFields));
            fields.Add(lineItem.properties.instanceData);

            return fields.ToArray();
        }

        public static object[] GetLineItemFieldsCsp(Item lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.usageStartTime);
            fields.Add(lineItem.usageEndTime);
            fields.Add(lineItem.resource.id);
            fields.Add(lineItem.resource.name);
            fields.Add(lineItem.resource.category);
            fields.Add(lineItem.resource.subcategory);
            fields.Add(lineItem.resource.region);
            fields.Add(lineItem.quantity);
            fields.Add(lineItem.unit);
            fields.Add(string.Empty); // Tags
            fields.Add(JsonUtils.ExtractInfoFields(lineItem.infoFields));
            fields.Add(lineItem.instanceData);
            fields.Add(lineItem.attributes);

            return fields.ToArray();
        }

        public static object[] GetLineItemFieldsEa(Datum lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.date);
            fields.Add(lineItem.accountId);
            fields.Add(lineItem.accountName);
            fields.Add(lineItem.productId);
            fields.Add(lineItem.product);
            fields.Add(lineItem.resourceLocationId);
            fields.Add(lineItem.resourceLocation);
            fields.Add(lineItem.consumedServiceId);
            fields.Add(lineItem.consumedService);
            fields.Add(lineItem.departmentId);
            fields.Add(lineItem.departmentName);
            fields.Add(lineItem.accountOwnerEmail);
            fields.Add(lineItem.serviceAdministratorId);
            fields.Add(lineItem.subscriptionId);
            fields.Add(lineItem.subscriptionGuid);
            fields.Add(lineItem.subscriptionName);
            fields.Add(lineItem.tags);
            fields.Add(lineItem.meterId);
            fields.Add(lineItem.meterName);
            fields.Add(lineItem.meterCategory);
            fields.Add(lineItem.meterSubCategory);
            fields.Add(lineItem.meterRegion);
            fields.Add(lineItem.consumedQuantity);
            fields.Add(lineItem.unitOfMeasure);
            fields.Add(lineItem.resourceRate);
            fields.Add(lineItem.Cost);
            fields.Add(lineItem.instanceId);
            fields.Add(lineItem.serviceInfo1);
            fields.Add(lineItem.serviceInfo2);
            fields.Add(lineItem.additionalInfo);
            fields.Add(lineItem.storeServiceIdentifier);
            fields.Add(lineItem.costCenter);
            fields.Add(lineItem.resourceGroup);

            return fields.ToArray();
        }

        public static object[] GetLineItemFields(Meter lineItem)
        {
            List<object> fields = new List<object>();
            var meterRate = ((string) ((Newtonsoft.Json.Linq.JProperty) lineItem.MeterRates.First).Value); // Value can be cast to double.
            fields.Add(lineItem.MeterId);
            fields.Add(lineItem.MeterName);
            fields.Add(lineItem.MeterCategory);
            fields.Add(lineItem.MeterSubCategory);
            fields.Add(lineItem.Unit);
            fields.Add(lineItem.MeterRegion);
            fields.Add(lineItem.MeterRates.Count > 0 ? meterRate : string.Empty);
            fields.Add(lineItem.EffectiveDate);
            fields.Add(lineItem.IncludedQuantity);

            return fields.ToArray();
        }
    }
}
