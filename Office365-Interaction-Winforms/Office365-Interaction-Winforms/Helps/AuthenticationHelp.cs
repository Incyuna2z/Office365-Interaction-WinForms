using System;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Configuration;

namespace Office365_Interaction_Winforms
{
    class AuthenticaionHelper
    {
        public static readonly string DiscoveryServiceResourceId = "https://api.office.com/discovery/";
        const string AuthorityFormat = "https://login.microsoftonline.com/{0}/";
        static readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        static readonly string ClientId = ConfigurationManager.AppSettings["ida:ClientId"].ToString();
        static string Domain = ConfigurationManager.AppSettings["ida:Domain"].ToString();
        static readonly Uri RedirectUri = new Uri("Your Redirect URI");

        static string TenantID = String.Empty;
        static string _authority = String.Empty;

        static AuthenticationContext authContext = null;

        public static string Authority
        {
            get
            {
                _authority = String.Format(AuthorityFormat, Domain);
                return _authority;
            }
        }

        public static async Task<AuthenticationResult> GetAccessToken(string serviceResourceId)
        {
            AuthenticationResult result = null;
            if (authContext == null)
            {
                authContext = new AuthenticationContext(Authority);

                result = await authContext.AcquireTokenAsync(serviceResourceId, ClientId, RedirectUri, new PlatformParameters(PromptBehavior.Always));
            }
            else
            {
                result = await authContext.AcquireTokenSilentAsync(serviceResourceId, ClientId);
            }

            TenantID = result.TenantId;

            return result;
        }
    }
}
