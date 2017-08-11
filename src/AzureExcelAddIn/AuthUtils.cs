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
        private const string AzureAuthUrl = "https://login.microsoftonline.com";
        private const string ApplicationId = "1950a258-227b-4e31-a9cf-717495945fc2";

        public static string GetAuthorizationHeader(string tenantId, bool forceReAuthentication)
        {
            var authUrl = string.Format(CultureInfo.InvariantCulture, "{0}/{1}", AzureAuthUrl, tenantId);
            var context = new AuthenticationContext(authUrl);
            
            // First try a silent auth.
            AuthenticationResult result;
            if (!forceReAuthentication)
            {
                try
                {
                    var userId = TokenCache.DefaultShared.ReadItems().FirstOrDefault();
                    if (userId != null)
                    {
                        result = context.AcquireTokenSilent(
                            AzureManagementResourceUrl,
                            ApplicationId,
                            new UserIdentifier(userId.UniqueId,
                                UserIdentifierType.OptionalDisplayableId));
                        return result.AccessToken;
                    }
                }
                catch { /* swallowing auth exceptions */ }
            }

            result = context.AcquireToken(
                AzureManagementResourceUrl,
                ApplicationId,
                new Uri(RedirectUrn),
                forceReAuthentication ? PromptBehavior.Always : PromptBehavior.Auto);
            return result.AccessToken;
        }
    }
}
