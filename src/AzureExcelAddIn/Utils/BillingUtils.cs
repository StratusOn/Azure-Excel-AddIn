using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelAddIn1
{
    internal class BillingUtils
    {
        public static async Task<Tuple<RateCard, string>> GetRateCardStandardAsync(string authorizationToken, string subscriptionId, string offerDurableId, string currency, string locale, string regionInfo, string apiVersion = "2015-06-01-preview")
        {
            string rateCardUrl =
            $"https://management.azure.com/subscriptions/{subscriptionId}/providers/Microsoft.Commerce/RateCard?api-version={apiVersion}&$filter=OfferDurableId eq '{offerDurableId}' and Currency eq '{currency}' and Locale eq '{locale}' and RegionInfo eq '{regionInfo}'";
            string content = await GetRestCallResultsAsync(authorizationToken, rateCardUrl);
            return new Tuple<RateCard, string>(JsonConvert.DeserializeObject<RateCard>(content), content);
        }

        public static async Task<Tuple<CspRateCard, string>> GetRateCardCspAsync(string authorizationToken, string currency, string locale, string regionInfo)
        {
            string rateCardUrl =
                $"https://api.partnercenter.microsoft.com/v1/ratecards/azure&currency={currency}&region={regionInfo}";
            var headers = new Dictionary<string, string>();
            headers.Add("X-Locale", locale);
            string content = await GetRestCallResultsAsync(authorizationToken, rateCardUrl, headers);
            return new Tuple<CspRateCard, string>(JsonConvert.DeserializeObject<CspRateCard>(content), content);
        }

        public static async Task<Tuple<PriceSheet, string>> GetPriceSheetAsync(string authorizationToken, string enrollmentNumber, string billingPeriod)
        {
            string rateCardUrl = string.IsNullOrWhiteSpace(billingPeriod) ?
                $"https://consumption.azure.com/v2/enrollments/{enrollmentNumber}/pricesheet" :
                $"https://consumption.azure.com/v2/enrollments/{enrollmentNumber}/billingPeriods/{billingPeriod}/pricesheet";
            string content = await GetRestCallResultsAsync(authorizationToken, rateCardUrl);
            return new Tuple<PriceSheet, string>(JsonConvert.DeserializeObject<PriceSheet>(content), content);
        }

        public static async Task<Tuple<UsageAggregates, string>> GetUsageAggregatesStandardAsync(string authorizationToken, string subscriptionId, string reportStartDate, string reportEndDate, string aggregationGranularity, string showDetails)
        {
            string usageAggregatesUrl =
                        $"https://management.azure.com/subscriptions/{subscriptionId}/providers/Microsoft.Commerce/UsageAggregates?api-version=2015-06-01-preview&reportedStartTime={reportStartDate}&reportedEndTime={reportEndDate}&aggregationGranularity={aggregationGranularity}&showDetails={showDetails}";
            string content = await GetRestCallResultsAsync(authorizationToken, usageAggregatesUrl);
            return new Tuple<UsageAggregates, string>(JsonConvert.DeserializeObject<UsageAggregates>(content), content);
        }

        public static async Task<Tuple<CspUsageAggregates, string>> GetUsageAggregatesCspAsync(string authorizationToken, string subscriptionId, string customerTenantId, string reportStartDate, string reportEndDate, string aggregationGranularity, string showDetails, int chunkSize)
        {
            string usageAggregatesUrl =
                $"https://api.partnercenter.microsoft.com/v1/customers/{customerTenantId}/subscriptions/{subscriptionId}/utilizations/azure?start_time={reportStartDate}&end_time={reportEndDate}&size={chunkSize}";
            string content = await GetRestCallResultsAsync(authorizationToken, usageAggregatesUrl);
            return new Tuple<CspUsageAggregates, string>(JsonConvert.DeserializeObject<CspUsageAggregates>(content), content);
        }

        public static async Task<Tuple<EaUsageAggregates, string>> GetUsageAggregatesEaAsync(string authorizationToken, string enrollmentNumber, string reportStartDate, string reportEndDate, string billingPeriod)
        {
            string usageAggregatesUrl;
            if (string.IsNullOrWhiteSpace(billingPeriod))
            {
                usageAggregatesUrl = $"https://consumption.azure.com/v2/enrollments/{enrollmentNumber}/usagedetailsbycustomdate?startTime={reportStartDate}&endTime={reportEndDate}";
            }
            else if (string.IsNullOrWhiteSpace(reportStartDate) || string.IsNullOrWhiteSpace(reportEndDate))
            {
                usageAggregatesUrl = $"https://consumption.azure.com/v2/enrollments/{enrollmentNumber}/{billingPeriod}/usagedetails";
            }
            else
            {
                usageAggregatesUrl = $"https://consumption.azure.com/v2/enrollments/{enrollmentNumber}/usagedetails";
            }

            string content = await GetRestCallResultsAsync(authorizationToken, usageAggregatesUrl);
            return new Tuple<EaUsageAggregates, string>(JsonConvert.DeserializeObject<EaUsageAggregates>(content), content);
        }

        public static async Task<string> GetRestCallResultsAsync(string authorizationToken, string url, IDictionary<string, string> headers = null)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            if (headers != null && headers.Count > 0)
            {
                foreach (var header in headers)
                {
                    request.Headers.Add(header.Key, header.Value);
                }
            }

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

            if (lineItem.properties.InstanceData != null)
            {
                if (lineItem.properties.InstanceData.MicrosoftResources.tags != null)
                {
                    fields.Add(string.Join(";", lineItem.properties.InstanceData.MicrosoftResources.tags.Select(x => x.Key + "=" + x.Value)));
                }
                else
                {
                    fields.Add(string.Empty);
                }

                fields.Add(lineItem.properties.InstanceData.MicrosoftResources.resourceUri);
                fields.Add(lineItem.properties.InstanceData.MicrosoftResources.location);
                fields.Add(lineItem.properties.InstanceData.MicrosoftResources.orderNumber);
                fields.Add(lineItem.properties.InstanceData.MicrosoftResources.partNumber);

                if (lineItem.properties.InstanceData.MicrosoftResources.additionalInfo != null)
                {
                    fields.Add(string.Join(";", lineItem.properties.InstanceData.MicrosoftResources.additionalInfo.Select(x => x.Key + "=" + x.Value)));
                }
                else
                {
                    fields.Add(string.Empty);
                }
            }
            else
            {
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
            }

            if (lineItem.properties.infoFields != null)
            {
                fields.Add(lineItem.properties.infoFields.meteredRegion);
                fields.Add(lineItem.properties.infoFields.meteredService);
                fields.Add(lineItem.properties.infoFields.meteredServiceType);
                fields.Add(lineItem.properties.infoFields.project);
                fields.Add(lineItem.properties.infoFields.serviceInfo1);
            }
            else
            {
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
            }

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

            if (lineItem.InstanceData != null)
            {
                if (lineItem.InstanceData.MicrosoftResources.tags != null)
                {
                    fields.Add(string.Join(";", lineItem.InstanceData.MicrosoftResources.tags.Select(x => x.Key + "=" + x.Value)));
                }
                else
                {
                    fields.Add(string.Empty);
                }

                fields.Add(lineItem.InstanceData.MicrosoftResources.resourceUri);
                fields.Add(lineItem.InstanceData.MicrosoftResources.location);
                fields.Add(lineItem.InstanceData.MicrosoftResources.orderNumber);
                fields.Add(lineItem.InstanceData.MicrosoftResources.partNumber);

                if (lineItem.InstanceData.MicrosoftResources.additionalInfo != null)
                {
                    fields.Add(string.Join(";", lineItem.InstanceData.MicrosoftResources.additionalInfo.Select(x => x.Key + "=" + x.Value)));
                }
                else
                {
                    fields.Add(string.Empty);
                }
            }
            else
            {
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
            }

            if (lineItem.infoFields != null)
            {
                fields.Add(lineItem.infoFields.meteredRegion);
                fields.Add(lineItem.infoFields.meteredService);
                fields.Add(lineItem.infoFields.meteredServiceType);
                fields.Add(lineItem.infoFields.project);
                fields.Add(lineItem.infoFields.serviceInfo1);
            }
            else
            {
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
                fields.Add(string.Empty);
            }

            fields.Add(lineItem.attributes.objectType);

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

        public static object[] GetRateCardLineItemFields(RateCard lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.Currency);
            fields.Add(lineItem.Locale);
            fields.Add(lineItem.MeterRegion);
            fields.Add(lineItem.IsTaxIncluded);

            if (lineItem.Tags != null)
            {
                fields.Add(string.Join(";", lineItem.Tags));
            }
            else
            {
                fields.Add(string.Empty);
            }

            return fields.ToArray();
        }

        public static object[] GetRateCardOfferTermLineItemFields(Offerterm lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.Name);
            fields.Add(lineItem.Credit);
            fields.Add(lineItem.EffectiveDate);

            if (lineItem.ExcludedMeterIds != null)
            {
                fields.Add(string.Join(";", lineItem.ExcludedMeterIds));
            }
            else
            {
                fields.Add(string.Empty);
            }

            if (lineItem.TieredDiscount != null)
            {
                fields.Add(string.Join(";", lineItem.TieredDiscount.Select(x => x.Key + "=" + x.Value)));
                fields.Add(lineItem.TieredDiscount.FirstOrDefault().Value);
            }
            else
            {
                fields.Add(string.Empty);
                fields.Add(string.Empty);
            }

            return fields.ToArray();
        }

        public static object[] GetRateCardMeterLineItemFields(Meter lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.MeterId);
            fields.Add(lineItem.MeterName);
            fields.Add(lineItem.MeterCategory);
            fields.Add(lineItem.MeterSubCategory);
            fields.Add(lineItem.Unit);
            fields.Add(lineItem.MeterRegion);

            if (lineItem.MeterRates != null)
            {
                fields.Add(string.Join(";", lineItem.MeterRates.Select(x => x.Key + "=" + x.Value)));
                fields.Add(lineItem.MeterRates.FirstOrDefault().Value);
            }
            else
            {
                fields.Add(string.Empty);
                fields.Add(string.Empty);
            }

            if (lineItem.MeterTags != null)
            {
                fields.Add(string.Join(";", lineItem.MeterTags));
            }
            else
            {
                fields.Add(string.Empty);
            }

            fields.Add(lineItem.EffectiveDate);
            fields.Add(lineItem.IncludedQuantity);
            fields.Add(lineItem.MeterStatus);

            return fields.ToArray();
        }

        public static object[] GetCspRateCardMeterLineItemFields(CspMeter lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.id);
            fields.Add(lineItem.name);
            fields.Add(lineItem.category);
            fields.Add(lineItem.subcategory);
            fields.Add(lineItem.unit);
            fields.Add(lineItem.region);

            if (lineItem.rates != null)
            {
                fields.Add(string.Join(";", lineItem.rates.Select(x => x.Key + "=" + x.Value)));
                fields.Add(lineItem.rates.FirstOrDefault().Value);
            }
            else
            {
                fields.Add(string.Empty);
                fields.Add(string.Empty);
            }

            if (lineItem.tags != null)
            {
                fields.Add(string.Join(";", lineItem.tags));
            }
            else
            {
                fields.Add(string.Empty);
            }

            fields.Add(lineItem.effectiveDate);
            fields.Add(lineItem.includedQuantity);
            fields.Add(lineItem.status);

            return fields.ToArray();
        }

        public static object[] GetPriceSheetMeterLineItemFields(PriceSheetMeter lineItem)
        {
            List<object> fields = new List<object>();
            fields.Add(lineItem.id);
            fields.Add(lineItem.billingPeriodId);
            fields.Add(lineItem.meterId);
            fields.Add(lineItem.meterName);
            fields.Add(lineItem.meterRegion);
            fields.Add(lineItem.unitOfMeasure);
            fields.Add(lineItem.includedQuantity);
            fields.Add(lineItem.partNumber);
            fields.Add(lineItem.unitPrice);
            fields.Add(lineItem.currencyCode);

            return fields.ToArray();
        }
    }
}
