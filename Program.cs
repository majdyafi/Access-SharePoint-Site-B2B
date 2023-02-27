using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace SharePoint_Access
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                RunAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }


        private static async Task RunAsync()
        {
            AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");

            IConfidentialClientApplication app;

            {
                app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                    .WithClientSecret(config.ClientSecret)
                    .WithAuthority(new Uri(config.Authority))
                    .Build();
            }

            
            app.AddInMemoryTokenCache();

            string[] scopes = new string[] { $"{config.ApiUrl}.default" }; // Generates a scope -> "https://graph.microsoft.com/.default"

            await CallSitesMSGraphUsingGraphSDK(app, scopes);
            
        }


        private static async Task CallSitesMSGraphUsingGraphSDK(IConfidentialClientApplication app, string[] scopes)
        {

            GraphServiceClient graphServiceClient = GetAuthenticatedGraphClient(app, scopes);

            List<Site> allUsers = new List<Site>();

            try
            {
                var site = await graphServiceClient
                    .Sites
                    .GetByPath("/sites/Knowledge", "majdio.sharepoint.com")
                    .Request()
                    .GetAsync();

                Console.WriteLine(site.Description);

            }
            catch (ServiceException e)
            {
                Console.WriteLine("Could not retrieve the Knowledge site: " + $"{e}");
            }

        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfidentialClientApplication app, string[] scopes)
        {            

            GraphServiceClient graphServiceClient =
                    new GraphServiceClient("https://graph.microsoft.com/V1.0/", new DelegateAuthenticationProvider(async (requestMessage) =>
                    {

                        AuthenticationResult result = await app.AcquireTokenForClient(scopes)
                            .ExecuteAsync();

                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));

            return graphServiceClient;
        }

        
    }
}
