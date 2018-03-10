using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
//using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;
//using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using System.Globalization;
using Microsoft.SharePoint.Client.Taxonomy;

using TaxonomyExtensions = Microsoft.SharePoint.Client.TaxonomyExtensions;

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using System.Xml.Serialization;

using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Utilities;
using Utility = Microsoft.SharePoint.Client.Utilities.Utility;
using System.Text.RegularExpressions;
using OfficeDevPnP.Core;
using System.Security;
using Microsoft.SharePoint.Client.Search.Query;

namespace Koskila.CitizenDeveloperTools
{
    /// <summary>
    /// Currently, does not work with lists with minor versions.
    /// </summary>
    public static class BetterCopyFunction
    {
        public static readonly string clientId = System.Configuration.ConfigurationManager.AppSettings["clientId"];
        public static readonly string clientSecret = System.Configuration.ConfigurationManager.AppSettings["clientSecret"];

        [FunctionName("BetterCopy2")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            string name = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "user", true) == 0)
                .Value;
            string pagelayout = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "pagelayout", true) == 0)
                .Value;


            // Get request body
            //string results = req.Content.ReadAsStringAsync().Result;

            dynamic data = await req.Content.ReadAsAsync<object>();

            // Set name to query string or body data
            name = name ?? data?.name.Email;

            string targetUrl = data?.targetUrl;
            string sourceUrl = data?.sourceUrl;
            Uri targetSiteUri = new Uri(targetUrl);
            Uri sourceSiteUri = new Uri(sourceUrl);
            string pageLayout = data?.pageLayout;


            int id = int.Parse((string) data?.targetId);

            // Get the realm for the URL
            var realm = TokenHelper.GetRealmFromTargetUrl(targetSiteUri);
            var tenantAdminUrl = ConfigurationManager.AppSettings["SiteCollectionRequests_TenantAdminSite"].TrimEnd(new[] { '/' });
            var tenantUrl = tenantAdminUrl.Substring(0, tenantAdminUrl.IndexOf(".com") + 4).Replace("-admin", "");
            AzureEnvironment env = TokenHelper.getAzureEnvironment(tenantAdminUrl);
            using (var ctx_target = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(targetSiteUri.ToString(), clientId, clientSecret, env))
            {
                using (var ctx_source = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(sourceSiteUri.ToString(), clientId, clientSecret, env))
                {
                    var exists = ctx_target.WebExistsFullUrl(targetUrl);

                    string targetWebUrl = targetUrl.Replace(tenantUrl, "");
                    var targetWeb = ctx_target.Site.OpenWeb(targetWebUrl);

                    ctx_target.Load(ctx_target.Site);
                    ctx_target.Load(targetWeb);
                    ctx_target.ExecuteQuery();

                    var sourceWeb = ctx_source.Site.OpenWeb(sourceUrl.Replace(tenantUrl, ""));
                    ctx_source.Load(sourceWeb);
                    ctx_source.ExecuteQuery();

                    log.Info(targetWeb.Title);

                    pageLayout = ctx_target.Site.Url + pageLayout.Substring(pageLayout.IndexOf("/_catalogs"));


                    var targetList = targetWeb.Lists.GetByTitle("Pages");
                    ctx_target.Load(targetList);
                    ctx_target.ExecuteQuery();

                    var sourceList = sourceWeb.Lists.GetByTitle("Pages");
                    ctx_source.Load(sourceList);
                    ctx_source.ExecuteQuery();

                    log.Info(targetList.Title + " " + targetList.ItemCount);

                    string publishingPageContent = "";

                    ListItem sourceItem = null;
                    try
                    {
                        string qs = String.Format("<View><Query><Where><Eq><FieldRef Name=\"ID\"></FieldRef><Value Type=\"Number\">{0}</Value></Eq></Where></Query></View>", id);
                        CamlQuery query = new CamlQuery();
                        query.ViewXml = qs;
                        var items = sourceList.GetItems(query);

                        ctx_source.Load(items);
                        ctx_source.ExecuteQuery();

                        sourceItem = items.First();
                    }
                    catch (Exception ex)
                    {
                        sourceItem = sourceWeb.GetListItem("/Pages/Forms/DispForm.aspx?ID=" + id);

                        //var items = sourceList.GetItems(CamlQuery.CreateAllItemsQuery());
                        ////var items = list.GetItems()
                        //ctx_source.Load(items);
                        //ctx_source.ExecuteQueryRetry();

                        //for (int i = 0; i < items.Count; i++)
                        //{
                        //    if (items[i].Id == id) sourceItem = items[i];
                        //}
                    }
                    finally
                    {
                        ctx_source.Load(sourceItem);
                        ctx_source.Load(sourceItem, r => r.Client_Title, r => r.Properties);
                        ctx_source.ExecuteQueryRetry();

                        log.Info(sourceItem.Client_Title);

                        publishingPageContent = sourceItem["PublishingPageContent"].ToString();
                    }

                    ListItem targetItem = null;
                    try
                    {
                        string qs = String.Format("<View><Query><Where><Eq><FieldRef Name=\"ID\"></FieldRef><Value Type=\"Number\">{0}</Value></Eq></Where></Query></View>", id);
                        CamlQuery query = new CamlQuery();
                        query.ViewXml = qs;
                        var items = targetList.GetItems(query);

                        ctx_target.Load(items);
                        ctx_target.ExecuteQuery();

                        targetItem = items.First();
                    }
                    catch (Exception ex)
                    {
                        //Thread.Sleep(1000 * 60);

                        targetItem = targetWeb.GetListItem("/Pages/Forms/DispForm.aspx?ID=" + id);

                        var items = targetList.GetItems(CamlQuery.CreateAllItemsQuery());
                        //var items = list.GetItems()
                        ctx_target.Load(items);
                        ctx_target.ExecuteQueryRetry();

                        for (int i = 0; i < items.Count; i++)
                        {
                            if (items[i].Id == id) targetItem = items[i];
                        }
                    }
                    finally
                    {
                        ctx_target.Load(targetItem);
                        ctx_target.Load(targetItem, r => r.Client_Title, r => r.Properties);
                        ctx_target.ExecuteQueryRetry();
                    }

                    log.Info(targetItem.Client_Title);

                    targetItem["PublishingPageLayout"] = pageLayout;
                    targetItem["PublishingPageContent"] = publishingPageContent;
                    targetItem.SystemUpdate();

                    ctx_target.ExecuteQuery();
                }
            }

            return name == null
            ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
            : req.CreateResponse(HttpStatusCode.OK, "Hello " + name);
        }
    }
}
