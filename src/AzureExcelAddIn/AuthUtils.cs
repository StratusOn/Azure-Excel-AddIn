using System;
using System.Globalization;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace ExcelAddIn1
{
    internal static class AuthUtils
    {
        private const string RedirectUrn = "urn:ietf:wg:oauth:2.0:oob";
        private const string AzureManagementResourceUrl = "https://management.core.windows.net/";
        private const string AzureAuthUrl = "https://login.microsoftonline.com";
        private const string ApplicationId = "1950a258-227b-4e31-a9cf-717495945fc2";

        public static string GetAuthorizationHeaderAsync(string tenantId, bool forceReAuthentication)
        {
            var authUrl = string.Format(CultureInfo.InvariantCulture, "{0}/{1}", AzureAuthUrl, tenantId);
            var context = new AuthenticationContext(authUrl);

            AuthenticationResult result = context.AcquireToken(
                AzureManagementResourceUrl, 
                ApplicationId, 
                new Uri(RedirectUrn), 
                forceReAuthentication ? PromptBehavior.Always : PromptBehavior.Auto);
            return result.AccessToken;
        }

    }
}
