using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace GraphTutorial
{
    public class SensitivityLabel
    {
        public string displayName
        {
            get;
            set;
        }
        public string id
        {
            get;
            set;
        }
        public bool protectionEnabled
        {
            get;
            set;
        }
    }

    public class GraphHelper
    {
        private static DeviceCodeCredential tokenCredential;
        private static GraphServiceClient graphClient;

        public static void Initialize(string clientId,
                                      string[] scopes,
                                      Func<DeviceCodeInfo, CancellationToken, Task> callBack)
        {
            tokenCredential = new DeviceCodeCredential(callBack, "", clientId);
            graphClient = new GraphServiceClient(tokenCredential, scopes);
        }

        public static async Task<string> GetAccessTokenAsync(string[] scopes)
        {
            var context = new TokenRequestContext(scopes);
            var response = await tokenCredential.GetTokenAsync(context);
            return response.Token;
        }

        public static async Task<User> GetMeAsync()
        {
            try
            {
                // GET /me
                return await graphClient.Me
                    .Request()
                    .Select(u => new
                    {
                        u.DisplayName,
                        u.MailboxSettings
                    })
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }

        public static async Task GetAllFilesAsync()
        {
            //IDriveItemChildrenCollectionPage driveItems = await graphClient.Me.Drive.Root.Children.Request().GetAsync();
            IDriveItemSearchCollectionPage driveItems = await graphClient.Me.Drive.Root.Search("").Request().GetAsync();
            List<string> driveItemIds = new List<string>();

            var driveItemsIterator = PageIterator<DriveItem>
            .CreatePageIterator(
            graphClient,
            driveItems,
            (item) =>
                {
                    if (item.File != null)
                    {
                        driveItemIds.Add(item.Id);
                    }
                    return true;
                },
            (req) =>
                {
                    return req;
                }
            );

            await driveItemsIterator.IterateAsync();

            foreach (var itemId in driveItemIds)
            {
                var driveItemInfo = await graphClient.Me.Drive.Items[itemId].Request().Select("Name,WebUrl,sensitivitylabel").GetAsync();
                object sensitivityLabelData;
                driveItemInfo.AdditionalData.TryGetValue("sensitivityLabel", out sensitivityLabelData);
                var sensitivityLabel = JsonSerializer.Deserialize<SensitivityLabel>(sensitivityLabelData.ToString());
                if(!string.IsNullOrEmpty(sensitivityLabel.id))
                {
                    Console.WriteLine("File Name: {0}, File Url: {1}, Sensitivitylabel Name: {2}, SensitivitylabelId: {3}, IsProtectectionEnabled: {4}",
                        driveItemInfo.Name, driveItemInfo.WebUrl, sensitivityLabel.displayName, sensitivityLabel.id, sensitivityLabel.protectionEnabled);
                }
            }
        }
    }
}