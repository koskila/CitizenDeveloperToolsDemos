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

            string siteUrl = data?.targetUrl;
            Uri siteUri = new Uri(siteUrl);
            string pageLayout = data?.pageLayout;

            int id = int.Parse((string) data?.targetId);

            // Get the realm for the URL
            var realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            var tenantAdminUrl = ConfigurationManager.AppSettings["SiteCollectionRequests_TenantAdminSite"].TrimEnd(new[] { '/' });
            var tenantUrl = tenantAdminUrl.Substring(0, tenantAdminUrl.IndexOf(".com") + 4).Replace("-admin", "");
            AzureEnvironment env = TokenHelper.getAzureEnvironment(tenantAdminUrl);
            using (var ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUri.ToString(), clientId, clientSecret, env))
            {
                var exists = ctx.WebExistsFullUrl(siteUrl);
                

                string weburl = siteUrl.Replace(tenantUrl, "");
                var web = ctx.Site.OpenWeb(weburl);

                ctx.Load(web);
                ctx.ExecuteQuery();

                log.Info(web.Title);

                var list = web.Lists.GetByTitle("Pages");

                ctx.Load(list);
                ctx.ExecuteQuery();

                log.Info(list.Title + " " + list.ItemCount);

                ctx.Load(ctx.Site);
                ctx.ExecuteQueryRetry();
                pageLayout = ctx.Site.Url + pageLayout.Substring(pageLayout.IndexOf("/_catalogs"));

                ListItem item = null;
                try
                {
                    string qs = String.Format("<View><Query><Where><Eq><FieldRef Name=\"ID\"></FieldRef><Value Type=\"Number\">{0}</Value></Eq></Where></Query></View>", id);
                    CamlQuery query = new CamlQuery();
                    query.ViewXml = qs;
                    var items = list.GetItems(query);

                    ctx.Load(items);
                    ctx.ExecuteQuery();

                    item = items.First();
                }
                catch (Exception ex)
                {
                    //Thread.Sleep(1000 * 60);

                    item = web.GetListItem("/Pages/Forms/DispForm.aspx?ID=" + id);

                    var items = list.GetItems(CamlQuery.CreateAllItemsQuery());
                    //var items = list.GetItems()
                    ctx.Load(items);
                    ctx.ExecuteQueryRetry();

                    for (int i = 0; i < items.Count; i++)
                    {
                        if (items[i].Id == id) item = items[i];
                    }
                }
                finally
                {
                    ctx.Load(item);
                    ctx.Load(item, r => r.Client_Title, r => r.Properties);
                    ctx.ExecuteQueryRetry();
                }

                log.Info(item.Client_Title);

                item["PublishingPageLayout"] = pageLayout;
                item.SystemUpdate();

                ctx.ExecuteQuery();
            }

            return name == null
            ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
            : req.CreateResponse(HttpStatusCode.OK, "Hello " + name);
        }
    }
}
