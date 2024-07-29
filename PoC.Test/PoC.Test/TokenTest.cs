using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Net.Http;

namespace PoC.Test
{
    public class TokenTest
    {
        private readonly HttpClient _client;

        public TokenTest()
        {
            _client = new HttpClient();
        }

        [FunctionName("TokenTest")]
        public async Task Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            var token = await GetSharepointToken();
        }

        public async Task<string> GetSharepointToken()
        {
            //_logger.LogInformation("Getting sharepoint Auth token");
            string access_Token = "";
            string strAuthURL = "";
            _client.DefaultRequestHeaders.Remove("authorization");
            _client.DefaultRequestHeaders.Add("authorization", "Bearer");
            _client.DefaultRequestHeaders.Add("cache-control", "no-cache");

            var parameters = new Dictionary<string, string>();
            parameters.Add("grant_type", "client_credentials");

            var SharepointclientId = Environment.GetEnvironmentVariable("SharepointclientId");
            var SharepointclientSecret = Environment.GetEnvironmentVariable("SharepointclientSecret");
            var SharePointResource = Environment.GetEnvironmentVariable("SharePointResource");

            parameters.Add("client_id", SharepointclientId);
            parameters.Add("client_secret", SharepointclientSecret);
            parameters.Add("resource", SharePointResource);

            var encodedContent = new FormUrlEncodedContent(parameters);
            strAuthURL = Environment.GetEnvironmentVariable("SharepointAuthURL");
            //_logger.LogInformation("Calling Sharepoint Auth service to get token, URL is " + strAuthURL);
            var httpResponse = await _client.PostAsync(strAuthURL, encodedContent);

            if (httpResponse.IsSuccessStatusCode)
            {
                var content = await httpResponse.Content.ReadAsStringAsync();
                var jsonResult = JsonConvert.DeserializeObject<JObject>(content);
                access_Token = (string)jsonResult["access_token"];
            }
            //_logger.LogInformation("Received Auth token from Sharepoint site");
            return access_Token;
        }
    }
}
