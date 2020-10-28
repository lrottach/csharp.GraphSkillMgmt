using System;
using System.Collections.Generic;
using GraphSkillMgmt.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace GraphSkillMgmt
{
    class Program
    {
        private static GraphServiceClient _graphClient;
        
        static void Main(string[] args)
        {
            // Load application config
            var config = LoadAppSettings();
            if (config == null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }
            
            // Initialize Graph client
            GraphServiceClient graphClient = GetAuthenticatedGraphClient(config);
            
            // Test query
            var graphRequest = graphClient.Users
                .Request()
                .Select(u => new 
                {
                    u.DisplayName, u.Mail
                })
                .Top(15)
                .Filter("startsWith(surname,'A') or startsWith(surname,'B') or startsWith(surname,'C')");

            var results = graphRequest.GetAsync().Result;
            foreach (var user in results)
            {
                Console.WriteLine($"{user.Id}: {user.DisplayName} <{user.Mail}>");
            }
            
            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(graphRequest.GetHttpRequestMessage().RequestUri);
            
            // Wait until confirmation
            Console.ReadLine();
        }

        static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                    .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", false, true)
                    .Build();

                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            // Get required variables from config
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            // This specific scope means, that the application will default to what is defined
            // in the application registration that using dynamic scopes
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");
            
            // Initialize ConfidentialClientApplicationBuilder
            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithAuthority(authority)
                .WithRedirectUri(redirectUri)
                .WithClientSecret(clientSecret)
                .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphClient = new GraphServiceClient(authenticationProvider);
            return _graphClient;
        }
    }
}