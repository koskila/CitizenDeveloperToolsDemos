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
            log.Info("C# HTTP trigger function BetterCopyFunction processed a request.");

            string errorMsg = "";

            try
            {
                dynamic data = await req.Content.ReadAsAsync<object>();

                string targetUrl = data?.targetUrl;
                string sourceUrl = data?.sourceUrl;
                Uri targetSiteUri = new Uri(targetUrl);
                Uri sourceSiteUri = new Uri(sourceUrl);
                string pageLayout = data?.pageLayout;

                string fileName = ""; // we get this from the source item

                int sourceId;
  
                try
                {
                    string strSourceId = data?.sourceId;
                    sourceId = int.Parse(strSourceId);
                }
                catch (Exception)
                {
                    log.Error("Setting up variables failed.");
                    errorMsg += "Setting up variables failed.";
                    throw;
                }

                log.Info("Got the variables! Now connecting to SharePoint...");

                // Get the realm for the URL
                var realm = TokenHelper.GetRealmFromTargetUrl(targetSiteUri);
                // parse tenant admin url from the sourceUrl (there's probably a cuter way to do this but this is simple :])
                string tenantAdminUrl = sourceUrl.Substring(0, sourceUrl.IndexOf(".com") + 4).TrimEnd(new[] { '/' }).Replace(".sharepoint", "-admin.sharepoint");
                // parse tenant url from the admin url
                var tenantUrl = tenantAdminUrl.Substring(0, tenantAdminUrl.IndexOf(".com") + 4).Replace("-admin", "");
                AzureEnvironment env = TokenHelper.getAzureEnvironment(tenantAdminUrl);

                using (var ctx_target = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(targetSiteUri.ToString(), clientId, clientSecret, env))
                {
                    log.Info("Target site context built successfully!");

                    using (var ctx_source = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(sourceSiteUri.ToString(), clientId, clientSecret, env))
                    {
                        log.Info("Source site context built successfully!");

                        var exists = ctx_target.WebExistsFullUrl(targetUrl);

                        string targetWebUrl = targetUrl.Replace(tenantUrl, "");
                        var targetWeb = ctx_target.Site.OpenWeb(targetWebUrl);

                        ctx_target.Load(ctx_target.Site);
                        ctx_target.Load(targetWeb);
                        ctx_target.ExecuteQuery();

                        var sourceWeb = ctx_source.Site.OpenWeb(sourceUrl.Replace(tenantUrl, ""));
                        ctx_source.Load(sourceWeb);
                        ctx_source.ExecuteQuery();

                        log.Info("SharePoint connection fine! Connected to: " + targetWeb.Title);

                        pageLayout = ctx_target.Site.Url + pageLayout.Substring(pageLayout.IndexOf("/_catalogs"));

                        var sourceList = sourceWeb.Lists.GetByTitle("Pages");
                        ctx_source.Load(sourceList);
                        ctx_source.ExecuteQuery();

                        var targetList = targetWeb.Lists.GetByTitle("Pages");
                        ctx_target.Load(targetList);
                        ctx_target.ExecuteQuery();

                        log.Info("... and: " + targetList.Title + " " + targetList.ItemCount);

                        string publishingPageContent = "";

                        ListItem sourceItem = null;
                        try
                        {
                            sourceItem = sourceList.GetItemById(sourceId);

                            if (sourceItem == null)
                            {
                                string qs = String.Format("<View><Query><Where><Eq><FieldRef Name=\"ID\"></FieldRef><Value Type=\"Number\">{0}</Value></Eq></Where></Query></View>", sourceId);
                                CamlQuery query = new CamlQuery();
                                query.ViewXml = qs;
                                var items = sourceList.GetItems(query);

                                ctx_source.Load(items);
                                ctx_source.ExecuteQuery();

                                sourceItem = items.First();
                            }
                        }
                        catch (Exception ex)
                        {
                            sourceItem = sourceWeb.GetListItem("/Pages/Forms/DispForm.aspx?ID=" + sourceId);
                            errorMsg += ex.Message + " ";
                        }
                        finally
                        {
                            ctx_source.Load(sourceItem);
                            ctx_source.Load(sourceItem.File);
                            ctx_source.Load(sourceItem, r => r.Client_Title, r => r.Properties);
                            ctx_source.ExecuteQueryRetry();

                            log.Info("Got source item! Title: " + sourceItem.Client_Title);

                            if (sourceItem["PublishingPageContent"] != null) publishingPageContent = sourceItem["PublishingPageContent"].ToString();
                        }

                        fileName = sourceItem.File.Name;

                        // at this point, we've fetched all the info we needed. On to getting the target item, and then updating the fields there.
                        ListItem targetItem = null;
                        try
                        {
                            targetItem = targetList.GetItemById(sourceId);

                            if (targetItem == null)
                            {
                                string qs1 = String.Format("<View><Query><Where><Eq><FieldRef Name=\"ID\"></FieldRef><Value Type=\"Number\">{0}</Value></Eq></Where></Query></View>", sourceId);
                                CamlQuery query1 = new CamlQuery();
                                query1.ViewXml = qs1;
                                var items1 = targetList.GetItems(query1);

                                ctx_target.Load(items1);
                                ctx_target.ExecuteQuery();

                                targetItem = items1.First();
                            }
                        }
                        catch (Exception ex)
                        {
                            log.Warning("Getting source item via conventional ways failed. Trying the unorthodox ones...");

                            targetItem = targetWeb.GetListItem("/Pages/Forms/DispForm.aspx?ID=" + sourceId);

                            var items = targetList.GetItems(CamlQuery.CreateAllItemsQuery());
                            ctx_target.Load(items);
                            ctx_target.ExecuteQueryRetry();

                            for (int i = 0; i < items.Count; i++)
                            {
                                if (items[i].Id == sourceId) targetItem = items[i];
                            }
                        }
                        finally
                        {
                            try
                            {
                                string str = "Published automatically by an Azure Function (BetterCopyFunction).";
                                targetItem.File.CheckIn(str, CheckinType.MajorCheckIn);
                                targetItem.File.Publish(str);

                                ctx_target.Load(targetItem);

                                ctx_target.ExecuteQueryRetry();
                            }
                            catch (Exception ex)
                            {
                                log.Info("Error: " + ex.Message);
                            }

                            ctx_target.Load(targetItem);
                            ctx_target.Load(targetItem, r => r.Client_Title, r => r.Properties);
                            ctx_target.ExecuteQueryRetry();
                        }

                        log.Info("Target item title: " + targetItem.Client_Title);

                        try
                        {
                            targetItem["PublishingPageLayout"] = pageLayout;
                            targetItem["PublishingPageContent"] = publishingPageContent;
                            targetItem.SystemUpdate();

                            ctx_target.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            log.Warning("There was an error in saving target item values. Values were: " + pageLayout + " " + publishingPageContent);
                            log.Warning("Error was: " + ex.Message);
                        }
                        finally
                        {
                            log.Info("Target item updated!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMsg += ex.Message;
                errorMsg += "\r\n " + ex.StackTrace;

                throw;
            }

            return String.IsNullOrEmpty(errorMsg)
                ? req.CreateResponse(HttpStatusCode.OK, "Function run was a success.")
                : req.CreateResponse(HttpStatusCode.InternalServerError, errorMsg);
        }
    }
}
