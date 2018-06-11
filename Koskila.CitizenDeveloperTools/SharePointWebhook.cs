using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;
using System.Configuration;
using System.Web.Script.Serialization;
using System.Text;

using Newtonsoft.Json;

using OfficeDevPnP.Core;

using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;

//using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;
//using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;

namespace Koskila.CitizenDeveloperTools
{
    public static class SharePointWebhook
    {
        private static readonly HttpClient _client = new HttpClient();
        private static readonly string _apiAddress = System.Configuration.ConfigurationManager.AppSettings["ApiAddress"];
        private static readonly string _notificationFlowAddress = System.Configuration.ConfigurationManager.AppSettings["NotificationFlowAddress"];

        public static readonly string clientId = System.Configuration.ConfigurationManager.AppSettings["clientId"];
        public static readonly string clientSecret = System.Configuration.ConfigurationManager.AppSettings["clientSecret"];

        private const bool enrichViaExternalAPI = false;

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

            // there should always be just one so let's get the first item
            var notification = notifications.First();

            // Get the realm for the URL
            var tenantAdminUrl = ConfigurationManager.AppSettings["SiteCollectionRequests_TenantAdminSite"].TrimEnd(new[] { '/' });
            var tenantUrl = tenantAdminUrl.Substring(0, tenantAdminUrl.IndexOf(".com") + 4).Replace("-admin", "");
            AzureEnvironment env = TokenHelper.getAzureEnvironment(tenantAdminUrl);
            log.Info($"Tenant url {tenantUrl} and notification from {notification.SiteUrl} ");

            string fullUrl = string.Format("{0}{1}", tenantUrl, notification.SiteUrl);
            log.Info($"{fullUrl}");
            Uri targetSiteUri = new Uri(fullUrl);
            log.Info($"Connecting to SharePoint at {targetSiteUri.AbsoluteUri}");

            var realm = TokenHelper.GetRealmFromTargetUrl(targetSiteUri);

            try
            {
                using (var ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(targetSiteUri.ToString(), clientId, clientSecret, env))
                {
                    log.Info("Connected to SharePoint!");

                    var ctxWeb = ctx.Site.OpenWebById(new Guid(notification.WebId));
                    try
                    {
                        ctx.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        log.Error("Error in ctx ExecuteQueryRetry, stage 1: " + ex.Message);
                        throw;
                    }

                    Guid listId = new Guid(notification.Resource);

                    List targetList = ctxWeb.Lists.GetById(listId);
                    ctx.Load(targetList, List => List.ParentWebUrl);
                    ctx.Load(targetList, List => List.Title);
                    ctx.Load(targetList, List => List.DefaultViewUrl);
                    ctx.Load(ctxWeb, Web => Web.Url);
                    ctx.ExecuteQueryRetry();

                    log.Info($"Got list {targetList.Title} at {ctxWeb.Url} !");

                    // now send the query to a custom API as a POST
                    var values = new Dictionary<string, string>();

                    if (notifications.Count > 0)
                    {
                        log.Info($"Processing notifications...");

                        StringContent stringcontent;
                        HttpResponseMessage apiresponse;

                        if (enrichViaExternalAPI)
                        {
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

                                values.Add("message" + i, m);

                                log.Info($"Notification {i} : {m}");
                            }

                            //stringcontent = new FormUrlEncodedContent(values);

                            apiresponse = await _client.PostAsync(_apiAddress, stringcontent);

                            var responseString = await apiresponse.Content.ReadAsStringAsync();

                            log.Info($"Got response: " + responseString);
                        }

                        // we have the response, now we let another flow know about it - through a call to the API!
                        var message = JsonConvert.SerializeObject(notification);
                        string link = tenantUrl + targetList.DefaultViewUrl;
                        var obj = new Dictionary<string, string>();
                        obj.Add("message", "New item on a list: " + targetList.Title);
                        obj.Add("link", link);

                        var serializer = new JavaScriptSerializer();
                        var json = serializer.Serialize(obj);
                        stringcontent = new StringContent(json, Encoding.UTF8, "application/json");

                        //stringcontent = new FormUrlEncodedContent(obj);

                        log.Info($"Now pushing this: " + stringcontent);
                        apiresponse = await _client.PostAsync(_notificationFlowAddress, stringcontent);

                        log.Info($"Pushed to Flow! Got this back: " + apiresponse);

                        // if we get here we assume the request was well received
                        return new HttpResponseMessage(HttpStatusCode.OK);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw;
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
