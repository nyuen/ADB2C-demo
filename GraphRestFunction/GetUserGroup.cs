// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <GetMyNewestMessageSnippet>
using GraphTutorial.Authentication;
using GraphTutorial.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel;
using System.Collections.Generic;

namespace GraphTutorial
{
   
   public struct Group {
       public string id;
       public string tenantType;
       public string name;
   }
    public class GetUserGroup
    {

        public GetUserGroup()
        {
           
        }

        [FunctionName("GetUserGroup")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req, ExecutionContext context, 
            ILogger log)
        {
            string objectId = req.Query["objectId"];

            var appConfig = new ConfigurationBuilder()
	                            .SetBasePath(context.FunctionAppDirectory)
            // This gives you access to your application settings in your local development environment
	                            .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true) 
            // This is what actually gets you the application settings in Azure
                                .AddEnvironmentVariables() 
                                .Build();
				
	   
            // Initialize the client credential auth provider
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(appConfig["AppId"])
                .WithTenantId(appConfig["TenantId"])
                .WithClientSecret(appConfig["ClientSecret"])
                .Build();
            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            // Set up the Microsoft Graph service client with client credentials
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

           var groupsRes = await graphClient
                    .Users[objectId]
                    //.MemberOf
                    .TransitiveMemberOf
                    .Request()
                    .GetAsync();
            
            var groups = new List<Group>();
            while (groupsRes.Count > 0)
            {
            foreach (var group in groupsRes)
            {
                
                if (group.ODataType.Equals("#microsoft.graph.group")){
                    var groupDetails = await graphClient.Groups[group.Id]
	                            .Request()
	                            .GetAsync();//Microsoft.Graph.Group
                    var currentGroup = new Group();
                    currentGroup.id = group.Id;
                    currentGroup.name = groupDetails.DisplayName;
                    currentGroup.tenantType = $"{groupDetails.AdditionalData["extension_275feff56d9b447c9ed3f7d79ec6236f_tenantType"]}";
                    groups.Add(currentGroup);
                }
            }
            
            if (groupsRes.NextPageRequest != null)
            {
                groupsRes = groupsRes.NextPageRequest.GetAsync().Result;
            }
            else
            {
                break;
            }
            }

            //return new OkObjectResult(
                return new JsonResult(
                new
                {
                    groups
                });
        }

    }
}


// </GetMyNewestMessageSnippet>
