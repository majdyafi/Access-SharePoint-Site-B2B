using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace SharePoint_Access;

public class GraphServiceClientInstance
{
  public static GraphServiceClient GetAuthenticatedGraphClient(IConfidentialClientApplication app, string[] scopes)
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