using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using System;
using System.Threading.Tasks;

namespace SharePoint_Access
{
  class Program
    {
        private const string SITE = "Knowledge";
        private const string HOST = "majdio.sharepoint.com";
        private const string MATERIAL = "Material";
        
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

            GraphServiceClient graphServiceClient = GraphServiceClientInstance.GetAuthenticatedGraphClient(app, scopes);
            
            var materials = await graphServiceClient
                .Sites
                .GetByPath($"sites/{SITE}", HOST)
                .Lists[MATERIAL]
                .Items
                .Request()
                .Expand(item => item.Fields)
                .GetAsync();

            foreach (var item in materials)
            {
                Console.WriteLine
                    ($"item: {item.Fields.AdditionalData["Title"]} {item.Fields.AdditionalData["URL"]} ");
            }
        }
        
    }
}
