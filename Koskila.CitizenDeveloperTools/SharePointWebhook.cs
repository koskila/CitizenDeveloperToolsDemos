using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System;
using Newtonsoft.Json;
using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using System.Collections.Generic;

namespace Koskila.CitizenDeveloperTools
{
    public static class SharePointWebhook
    {
        private static readonly HttpClient _client = new HttpClient();
        private static readonly string _apiAddress = System.Configuration.ConfigurationManager.AppSettings["ApiAddress"];
        private static readonly string _notificationFlowAddress = System.Configuration.ConfigurationManager.AppSettings["NotificationFlowAddress"];

        /// <summary>
        /// https://docs.microsoft.com/en-us/sharepoint/dev/apis/webhooks/sharepoint-webhooks-using-azure-functions
        /// </summary>
        /// <param name="req"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        [FunctionName("SharePointWebhook")]
        public static async Task<object> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info($"Webhook was triggered!");

            // Grab the validationToken URL parameter
            string validationToken = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                .Value;

            // If a validation token is present, we need to respond within 5 seconds by  
            // returning the given validation token. This only happens when a new 
            // web hook is being added
            if (validationToken != null)
            {
                log.Info($"Validation token {validationToken} received");
                var response = req.CreateResponse(HttpStatusCode.OK);
                response.Content = new StringContent(validationToken);
                return response;
            }

            log.Info($"SharePoint triggered our webhook...great :-)");
            var content = await req.Content.ReadAsStringAsync();
            log.Info($"Received following payload: {content}");

            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;
            log.Info($"Found {notifications.Count} notifications");

            // now send the query to a custom API as a POST
            var values = new Dictionary<string, string>();

            if (notifications.Count > 0)
            {
                log.Info($"Processing notifications...");
                for (int i = 0; i < notifications.Count; i++)
                {
                    var n = notifications[i];
                    //        CloudStorageAccount storageAccount = CloudStorageAccount.Parse("<YOUR STORAGE ACCOUNT>");
                    //        // Get queue... create if does not exist.
                    //        CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                    //        CloudQueue queue = queueClient.GetQueueReference("sharepointlistwebhookeventazuread");
                    //        queue.CreateIfNotExists();

                    //        // add message to the queue
                    string m = JsonConvert.SerializeObject(n);
                    //        log.Info($"Before adding a message to the queue. Message content: {message}");
                    //        queue.AddMessage(new CloudQueueMessage(message));
                    //        log.Info($"Message added :-)");

                    values.Add("message"+i, m);
                }

                var stringcontent = new FormUrlEncodedContent(values);

                var apiresponse = await _client.PostAsync(_apiAddress, stringcontent);

                var responseString = await apiresponse.Content.ReadAsStringAsync();

                log.Info($"Got this: " + responseString);

                // we have the response, now we let another flow know about it - through a call to the API!
                var notification = notifications.First();
                var message = JsonConvert.SerializeObject(notification);
                string link = notification.SiteUrl;
                var obj = new Dictionary<string, string>();
                obj.Add("message", "Webhook triggered! Message: " + message);
                obj.Add("link", link);
                stringcontent = new FormUrlEncodedContent(obj);

                log.Info($"Now pushing this: " + message);
                apiresponse = await _client.PostAsync(_notificationFlowAddress, stringcontent);

                log.Info($"Pushed to Flow! Got this back: " + apiresponse);

                // if we get here we assume the request was well received
                return new HttpResponseMessage(HttpStatusCode.OK);
            }

            log.Info($"Got nothing! Logging bad request.");
            return new HttpResponseMessage(HttpStatusCode.BadRequest);
        }

        // supporting classes
        public class ResponseModel<T>
        {
            [JsonProperty(PropertyName = "value")]
            public List<T> Value { get; set; }
        }

        public class NotificationModel
        {
            [JsonProperty(PropertyName = "subscriptionId")]
            public string SubscriptionId { get; set; }

            [JsonProperty(PropertyName = "clientState")]
            public string ClientState { get; set; }

            [JsonProperty(PropertyName = "expirationDateTime")]
            public DateTime ExpirationDateTime { get; set; }

            [JsonProperty(PropertyName = "resource")]
            public string Resource { get; set; }

            [JsonProperty(PropertyName = "tenantId")]
            public string TenantId { get; set; }

            [JsonProperty(PropertyName = "siteUrl")]
            public string SiteUrl { get; set; }

            [JsonProperty(PropertyName = "webId")]
            public string WebId { get; set; }
        }

        public class SubscriptionModel
        {
            [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
            public string Id { get; set; }

            [JsonProperty(PropertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
            public string ClientState { get; set; }

            [JsonProperty(PropertyName = "expirationDateTime")]
            public DateTime ExpirationDateTime { get; set; }

            [JsonProperty(PropertyName = "notificationUrl")]
            public string NotificationUrl { get; set; }

            [JsonProperty(PropertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
            public string Resource { get; set; }
        }
    }
}
