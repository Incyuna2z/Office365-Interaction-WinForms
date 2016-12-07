using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.SharePoint.CoreServices;

namespace Office365_Interaction_Winforms
{
    class O365Help
    {
        private static IDictionary<string, CapabilityDiscoveryResult> capabilities = null;

        private static async Task getCapabilities()
        {
            DiscoveryClient discoveryClient = new DiscoveryClient(
                async () =>
                {
                    var authResult = await AuthenticaionHelper.GetAccessToken(AuthenticaionHelper.DiscoveryServiceResourceId);
                    return authResult.AccessToken;
                });
            capabilities = await discoveryClient.DiscoverCapabilitiesAsync();
        }

        public async static Task<SharePointClient> CreateSharePointClientAsync(string capability)
        {
            if (capabilities == null)
            {
                await getCapabilities();
            }
            var myCapability = capabilities
                                        .Where(s => s.Key == capability)
                                        .Select(p => new { Key = p.Key, ServiceResourceId = p.Value.ServiceResourceId, ServiceEndPointUri = p.Value.ServiceEndpointUri })
                                        .FirstOrDefault();
            SharePointClient spClient = new SharePointClient(myCapability.ServiceEndPointUri,
                     async () =>
                     {
                         var authResult = await AuthenticaionHelper.GetAccessToken(myCapability.ServiceResourceId);
                         return authResult.AccessToken;
                     });
            return spClient;
        }

        public static async Task<List<Microsoft.Office365.SharePoint.FileServices.IItem>> getMyFiles(SharePointClient spClient)
        {
            List<Microsoft.Office365.SharePoint.FileServices.IItem> myFiles = new List<Microsoft.Office365.SharePoint.FileServices.IItem>();
            var myFileResult = await spClient.Files.ExecuteAsync();
            do
            {
                var files = myFileResult.CurrentPage;
                foreach (var file in files)
                {
                    myFiles.Add(file);
                }
                myFileResult = await myFileResult.GetNextPageAsync();
            } while (myFileResult != null);
            return myFiles;
        }
    }
}
