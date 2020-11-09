using OutlookWebAddIn2Web.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Net.Http.Headers;

namespace OutlookWebAddIn2Web.Controllers
{
    [Authorize] 
    public class WebController : ApiController
    {
      private IEnumerable<Customer> data = new List<Customer>() { new Customer("4711", "Customer 1"), 
                                                                  new Customer("4712", "Customer 2"), 
                                                                  new Customer("4713", "Customer 3" ), 
                                                                  new Customer("4714", "Customer 4"), 
                                                                  new Customer("4715", "Customer 5"), 
                                                                  new Customer("4716", "Customer 6"), 
                                                                  new Customer("4717", "Customer 7"), 
                                                                  new Customer("4718", "Customer 8"), 
                                                                  new Customer("4719", "Customer 9"), 
                                                                  new Customer("4720", "Customer 10"), 
                                                                  new Customer("4721", "Customer 11"), 
                                                                  new Customer("4722", "Customer 12") };

    // api/<controller>/Search/<querystring>
    [HttpGet]
    public IEnumerable<Customer> Search(string query)
    {
      return data.Where(c => c.Name.Contains(query));
    }

    // api/<controller>/GetMimeMessage
    [HttpPost]
    public async Task<IHttpActionResult> GetMimeMessage([FromBody] MimeMail request)
    {      
      if (Request.Headers.Contains("Authorization") && request.IsValid())
      {
        string accessToken = await this.GetAccessToken();
        if (accessToken.StartsWith("Error"))
        {
          return BadRequest(accessToken);
        }
        else
        {
          string mimeMailContent = await this.GetMime(accessToken, request.MessageID);
          return Content(HttpStatusCode.OK, mimeMailContent);
        }        
      }
      else
      {
        return BadRequest("Authentication header or parameter is missing.");
      }
    }

    private async Task<string> GetAccessToken()
    {
      var scopes = ClaimsPrincipal.Current.Claims;
      var identities = ClaimsPrincipal.Current.Identities;
      var scopeClaim = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope");
      if (scopeClaim != null)
      {
        // Check the allowed scopes
        string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
        if (!addinScopes.Contains("access_as_user"))
        {
          var msg = new HttpResponseMessage(HttpStatusCode.Unauthorized) { ReasonPhrase = "Missing access_as_user." };
          throw new HttpResponseException(msg);
        }
      }
      else
      {
        return "Error: The bearer token is invalid.";
      }

      string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
      UserAssertion userAssertion = new UserAssertion(bootstrapContext);

      string authority = String.Format(ConfigurationManager.AppSettings["Authority"], ConfigurationManager.AppSettings["DirectoryID"]);


      var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ClientID"])
                                                     .WithRedirectUri("https://localhost:44384")
                                                     .WithClientSecret(ConfigurationManager.AppSettings["ClientSecret"])
                                                     .WithAuthority(authority)
                                                     .Build();

      string[] graphScopes = { "https://graph.microsoft.com/Mail.Read" };
      AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
      AuthenticationResult authResult = null;
      string mimeMailContent = String.Empty;
      try
      {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
        return authResult.AccessToken;
      }
      catch (MsalServiceException e)
      {
        // multi-factor authentication.
        if (e.Message.StartsWith("AADSTS50076"))
        {
          string responseMessage = String.Format("Error: {{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
          return responseMessage;
        }
        // Lack of consent and invalid scope (permission).
        if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
        {
          return  String.Format("Error: Forbidden {0}", e.Message);
        }
        // All other MsalServiceExceptions.
        return String.Format("Error while exchanging to access token", e.Message);
      }
  }
  public Customer Get(string id)
    {
      return data.FirstOrDefault(c => c.ID == id);
    }

    // GET api/<controller>
    public IEnumerable<Customer> Get()
    {
      return data;
    }

    private async Task<string> GetMime(string accessToken, string mailID)
    {
      GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
          async (requestMessage) =>
          {
            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
          }));

      var mimeContent = await graphClient.Me.Messages[mailID]
        .Content
        .Request()        
        .GetAsync();
      string mimeContentStr = string.Empty;
      using (var Reader = new System.IO.StreamReader(mimeContent))
      {
        mimeContentStr = Reader.ReadToEnd();
      }
      return mimeContentStr;
    }
  }
}
