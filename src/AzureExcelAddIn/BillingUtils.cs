using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelAddIn1
{
    internal class BillingUtils
    {
        public static async Task<UsageAggregates> GetUsageAggregates(string authorizationToken, string subscriptionId, string reportStartDate, string reportEndDate, string aggregationGranularity, string showDetails)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);
            HttpResponseMessage response;

            string usageAggregatesUrl =
                $"https://management.azure.com/subscriptions/{subscriptionId}/providers/Microsoft.Commerce/UsageAggregates?api-version=2015-06-01-preview&reportedStartTime={reportStartDate}T00%3a00%3a00%2b00%3a00&reportedEndTime={reportEndDate}T00%3a00%3a00%2b00%3a00&aggregationGranularity={aggregationGranularity}&showDetails={showDetails}";

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, usageAggregatesUrl);
            //request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            response = await client.SendAsync(request);
            //string responseBody = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                string content = await response.Content.ReadAsStringAsync();
                var usageAggregatesResponse = JsonConvert.DeserializeObject<UsageAggregates>(content);

                return usageAggregatesResponse;
            }

            return null;
        }

        public static async Task<UsageAggregates> GetUsageAggregates(string authorizationToken, string nextLink)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authorizationToken);
            HttpResponseMessage response;

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, nextLink);
            //request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            response = await client.SendAsync(request);
            //string responseBody = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                string content = await response.Content.ReadAsStringAsync();
                var usageAggregatesResponse = JsonConvert.DeserializeObject<UsageAggregates>(content);

                return usageAggregatesResponse;
            }

            return null;
        }

    }
}
