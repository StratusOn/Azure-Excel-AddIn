using System;
using System.Globalization;
using System.Linq;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace ExcelAddIn1
{
    internal static class AuthUtils
    {
        private const string RedirectUrn = "urn:ietf:wg:oauth:2.0:oob";
        private const string AzureManagementResourceUrl = "https://management.core.windows.net/";
        private const string PartnerServiceResourceUrl = "https://api.partnercenter.microsoft.com";
        private const string AzureAuthUrl = "https://login.microsoftonline.com";
        private const string ApplicationId = "1950a258-227b-4e31-a9cf-717495945fc2";

        public static string GetAuthorizationHeader(string tenantId, bool forceReAuthentication, UsageApi usageApi)
        {
            return GetAuthorizationHeader(tenantId, forceReAuthentication, usageApi, null, null);
        }

        public static string GetAuthorizationHeader(string tenantId, bool forceReAuthentication, UsageApi usageApi, string customApplicationId, string customApplicationKey)
        {
            var authUrl = string.Format(CultureInfo.InvariantCulture, "{0}/{1}", AzureAuthUrl, tenantId);
            var context = new AuthenticationContext(authUrl);
            var resourceUrl = usageApi == UsageApi.CloudSolutionProvider
                ? PartnerServiceResourceUrl
                : AzureManagementResourceUrl;
            var customApplicationIdSpecified = !string.IsNullOrWhiteSpace(customApplicationId);
            var applicationId = customApplicationIdSpecified ? customApplicationId : ApplicationId;

            ClientCredential clientCredential = null;
            if (customApplicationIdSpecified && !string.IsNullOrWhiteSpace(customApplicationKey))
            {
                clientCredential = new ClientCredential(applicationId, customApplicationKey);
            }

            AuthenticationResult result;
            if (!forceReAuthentication)
            {
                // First try a silent auth.
                try
                {
                    var userId = TokenCache.DefaultShared.ReadItems().FirstOrDefault();
                    if (userId != null)
                    {
                        if (clientCredential != null)
                        {
                            result = context.AcquireTokenSilent(
                                resourceUrl,
                                clientCredential,
                                new UserIdentifier(userId.UniqueId,
                                    UserIdentifierType.OptionalDisplayableId));
                        }
                        else
                        {
                            result = context.AcquireTokenSilent(
                                resourceUrl,
                                applicationId,
                                new UserIdentifier(userId.UniqueId,
                                    UserIdentifierType.OptionalDisplayableId));
                        }

                        return result.AccessToken;
                    }
                }
                catch { /* swallowing auth exceptions only for silent auth */ }
            }

            if (clientCredential != null)
            {
                result = context.AcquireToken(
                    resourceUrl,
                    clientCredential);
            }
            else
            {
                result = context.AcquireToken(
                    resourceUrl,
                    applicationId,
                    new Uri(RedirectUrn),
                    forceReAuthentication ? PromptBehavior.Always : PromptBehavior.Auto);
            }

            return result.AccessToken;
        }
    }
}
