// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace Microsoft.PowerToys.Run.Plugin.Graph
{
    internal static class GraphHelper
    {
        public static GraphServiceClient GetGraphClient(string clientId, IEnumerable<string> scopes)
        {
            var pca = PublicClientApplicationBuilder.Create(clientId).Build();
            TokenCacheHelper.EnableSerialization(pca.UserTokenCache);

            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async requestMessage =>
            {
                var result = await GetAuthenticationResultAsync(pca, scopes);
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
            }));

            return graphClient;
        }

        private static async Task<AuthenticationResult> GetAuthenticationResultAsync(IPublicClientApplication pca, IEnumerable<string> scopes)
        {
            try
            {
                var accounts = await pca.GetAccountsAsync();
                return await pca.AcquireTokenSilent(scopes, accounts.FirstOrDefault()).ExecuteAsync();
            }
            catch (MsalUiRequiredException)
            {
                if (pca is IPublicClientApplication publicApp)
                {
                    return await publicApp.AcquireTokenInteractive(scopes).WithUseEmbeddedWebView(false).ExecuteAsync();
                }
                else
                {
                    throw;
                }
            }
        }
    }
}
