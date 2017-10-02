using System.Collections.Generic;

namespace ExcelAddIn1
{
    public class PersistedData
    {
        public string TenantId { get; set; }

        public string SubscriptionId { get; set; }

        public string EnrollmentNumber { get; set; }

        public string EaApiKey { get; set; }

        public string ApplicationId { get; set; }

        public string ApplicationKey { get; set; }

        public string CustomerTenantId { get; set; }

        public List<string> TenantIds { get; set; }

        public List<string> SubscriptionIds { get; set; }

        public List<string> EnrollmentNumbers { get; set; }

        public List<string> ApplicationIds { get; set; }

        public List<string> ApplicationKeys { get; set; }

        public List<string> CustomerTenantIds { get; set; }
    }
}
