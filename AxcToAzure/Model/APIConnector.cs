using AxcToAzure.Model;
using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using Resx = AxcToAzure.Properties.Resources;

namespace AxcToAzure.Model
{
  public class APIConnector
  {

    #region Properties
    private HttpClientHandler handler = new();
    private HttpClient client = new();
    public string Url { get; set; }
    public bool Initialized { get; set; }
    public object JsonConvert { get; private set; }
    private string createBody = @"{
""updatePackage"":""[{\""id\"":0,\""rev\"":0,\""projectId\"":\""\"",\""isDirty\"":true,\""tempId\"":-1,\""fields\"":{\""1\"":\""itemName\"",\""2\"":\""New\"",\""22\"":\""New\"",employeeValue\""25\"":\""itemType\"",\""10007\"":{\""type\"":1},\""10015\"":2,itemValue\""-2\"":apiTeamId,\""-104\"":apiProjectId}itemLink}]""
}";
    private string addedLink = @",\""links\"":{\""addedLinks\"":[{\""ID\"":parentId,\""LinkType\"":-2,\""Comment\"":\""\"",\""FldID\"":37,\""Changed Date\"":\""\\/azDate\\/\"",\""Revised Date\"":\""\\/azDate\\/\"",\""isAddedBySystem\"":true}]}";
    private string itemValue = @"\""10018\"":\""Business\"",";
    private string getBody = @"{""contributionIds"":[""ms.vss-work-web.backlogs-hub-backlog-data-provider""],""context"":{""properties"":{""forecasting"":false,""inProgress"":true,""completedChildItems"":true,""pageSource"":{""contributionPaths"":[""VSS"",""VSS/Resources"",""q"",""knockout"",""mousetrap"",""mustache"",""react"",""react-dom"",""react-transition-group"",""jQueryUI"",""jquery"",""OfficeFabric"",""tslib"",""@uifabric"",""VSSUI"",""ContentRendering"",""ContentRendering/Resources"",""WidgetComponents"",""WidgetComponents/Resources"",""TFSUI"",""TFSUI/Resources"",""Charts"",""Charts/Resources"",""TFS"",""Notifications"",""Presentation/Scripts/marked"",""Presentation/Scripts/URI"",""Presentation/Scripts/punycode"",""Presentation/Scripts/IPv6"",""Presentation/Scripts/SecondLevelDomains"",""highcharts"",""highcharts.more"",""highcharts.accessibility"",""highcharts.heatmap"",""highcharts.funnel"",""Analytics""],""diagnostics"":{""sessionId"":"""",""activityId"":"""",""bundlingEnabled"":true,""webPlatformVersion"":""M153"",""serviceVersion"":""Dev17.M153.5 (build: unknown)""},""navigation"":{""topMostLevel"":8,""area"":"""",""currentController"":""Apps"",""currentAction"":""ContributedHub"",""commandName"":""agile.backlogs-content"",""routeId"":""ms.vss-work-web.backlogs-content-route"",""routeTemplates"":[""{project}/_backlogs/{pivot}/{teamName}/{backlogLevel}"",""{project}/_backlogs/{pivot}/{teamName}"",""{project}/_backlogs/{pivot}""],""routeValues"":{""controller"":""Apps"",""action"":""ContributedHub"",""project"":""getProjectName"",""teamName"":""getTeamName"",""pivot"":""backlog"",""viewname"":""content"",""backlogLevel"":""Epics""}},""project"":{""id"":""getProjectId"",""name"":""getProjectName""},""selectedHubGroupId"":""ms.vss-work-web.work-hub-group"",""selectedHubId"":""ms.vss-work-web.backlogs-hub"",""url"":""getUrl""},""sourcePage"":{""url"":""getUrl"",""routeId"":""ms.vss-work-web.backlogs-content-route"",""routeValues"":{""controller"":""Apps"",""action"":""ContributedHub"",""project"":""getProjectName"",""teamName"":""getTeamName"",""pivot"":""backlog"",""viewname"":""content"",""backlogLevel"":""Epics""}}}}}
";
    private string getUrl = "";
    private string updateBody = @"{""updatePackage"":""[{\""id\"":azureId,\""rev\"":revision,\""projectId\"":\""scopeValue\"",\""isDirty\"":true,\""fields\"":{\""1\"":\""itemName\""}}]""}";
    private string teamId = "";
    private string projectId ="";
    public List<string> ErrorItems { get; set; }
    public List<DataItem> OnlineBacklog { get; set; }
    private string scopeValue { get; set; }
    //public Dictionary<int, WorkProject> WorkProjects = new();
    #endregion Properties

    #region Construktor
    public APIConnector(string username, SecureString password, string url, string apiteamId ="", string apiprojectId ="",string proxyAdress = "")
    {
      ErrorItems = new List<string>();
      if (proxyAdress != "")
      {

        WebProxy proxy = new WebProxy
        {
          Address = new Uri(proxyAdress),
          BypassProxyOnLocal = false,
          UseDefaultCredentials = false,
        };
        // Create a client handler that uses the proxy
        handler.Proxy = proxy;
        // Disable SSL verification
        handler.ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;
      }
      // Set Username and password
      var credential = new NetworkCredential(username, password);//, domain);

        teamId = apiteamId == null ? "" : apiteamId;
        projectId = apiprojectId == null ? "" :apiprojectId ;
      
      handler.Credentials = credential;
      // Finally, create the HTTP client object
      client = new HttpClient(handler: handler, disposeHandler: true);
      Url = url;
      Initialized = false;
    }

    #endregion Construktor

    #region Methods
    public async Task<bool> InitializeConnection()
    {
      try
      {
        //Aus angegebener URL Teamnamen schneiden
        int startindex = Url.LastIndexOf("/backlog/") + 9;
        string categoryName = (Url.Contains("/Epics")) ? "/Epics" : (Url.Contains("/Features")) ? "/Features" : (Url.Contains("/Stories")) ? "/Stories" : "/"; 
        var teamName = Url.Substring(startindex, Url.LastIndexOf(categoryName) - startindex).Replace("%20", " ");
        // Cut till Projectname
        startindex = Url.LastIndexOf("/tfs/") + 5;
        var shortenedUrl = Url.Substring(startindex);
        startindex = shortenedUrl.IndexOf("/") + 1;
        shortenedUrl = shortenedUrl.Substring(startindex);
        int endindex = shortenedUrl.IndexOf("/");
        var projectName = shortenedUrl.Substring(0, endindex);
        getUrl = Url.Substring(0, Url.IndexOf("/" + projectName));
        var apiUrl = Url.Substring(0, Url.LastIndexOf("/_backlogs/")) + "?__rt=fps";
        //Finde scopeValue des Projekts für ApiUrl
        var result = await BaseGetRequestAsync(apiUrl);
        var response = await result.Content.ReadAsStringAsync();
        if (!result.IsSuccessStatusCode) { throw new Exception(result.ReasonPhrase + "\n" + response); }
        var jsonResponse = JObject.Parse(response);
        scopeValue = jsonResponse["fps"].Value<JObject>("dataProviders").Value<string>("scopeValue");
        apiUrl = apiUrl.Substring(0, apiUrl.LastIndexOf("/")) + "/" + scopeValue + "/_api/_wit/nodes?__v=5";
        //Finde Projekt und Team id
        result = await BaseGetRequestAsync(apiUrl);
        response = await result.Content.ReadAsStringAsync();
        if (!result.IsSuccessStatusCode) { throw new Exception(result.ReasonPhrase + "\n" + response); }
        jsonResponse = JObject.Parse(response);
        if(teamId.Trim() == "" ||  projectId.Trim() == "")
        {

        projectId = "";
        teamId = "";
        var teams = jsonResponse.Value<JArray>("children")[0].Value<JArray>("children"); 
        var projects = jsonResponse.Value<JArray>("children")[1].Value<JArray>("children");
        var potentialProjectId = projects.Where(x => x.Value<string>("name") == teamName).FirstOrDefault();
        var potentialTeamId = teams.Where(x => x.Value<string>("name") == teamName).FirstOrDefault();
        if (potentialTeamId != null) { 
          teamId = potentialTeamId.Value<string>("id"); 
          projectId = (potentialProjectId != null) ? potentialProjectId.Value<string>("id"):jsonResponse.Value<string>("id");
        }
        else
        {
          teams = jsonResponse.Value<JArray>("children")[0].Value<JArray>("children");
          bool foundBacklog = false;
          foreach (var potTeam in teams)
          {
            var potName = potTeam.Value<string>("name");
            var Result = MessageBox.Show(Resx.ApiConnectorNameNotFound+$" {potName}", Resx.MessageAlert, MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (Result == MessageBoxResult.Yes)
            {
              potentialProjectId = projects.Where(x => x.Value<string>("name") == potName).FirstOrDefault();
              projectId = (potentialProjectId != null) ? potentialProjectId.Value<string>("id") : jsonResponse.Value<string>("id");
              teamId = potTeam.Value<string>("id");
              foundBacklog = true;
              break;


            }
          }
          if (!foundBacklog)
          {
            throw new Exception(Resx.ApiConnectorBacklogNotFound);
          }
        }
        }




        // Post Body vorbereiten
        createBody = createBody.Replace("apiTeamId", teamId).Replace("apiProjectId", projectId);
        getBody = getBody.Replace("getProjectName", projectName).Replace("getTeamName", teamName).Replace("getProjectId", scopeValue).Replace("getUrl", Url);
        Initialized = true;
        return true;
      }
      catch (Exception ex) { MessageBox.Show(ex.Message, Resx.MessageError, MessageBoxButton.OK, MessageBoxImage.Error); return false; }
    }
    public async Task<bool> GetExistingBacklog()
    {
      try
      {
        OnlineBacklog = new List<DataItem>();
        string body = getBody;
        //Post Api Url bereitmachen
        string apiUrl = getUrl + "/_apis/Contribution/dataProviders/query";
        client.DefaultRequestHeaders.Add("Accept", "application/json;api-version=5.1-preview.1");
        var result = await BasePostRequestAsync(apiUrl, body);
        client.DefaultRequestHeaders.Clear();
        var response = await result.Content.ReadAsStringAsync();
        if (!result.IsSuccessStatusCode) { throw new Exception(result.ReasonPhrase + "\n" + response); }
        var jsonResponse = JObject.Parse(response);
        jsonResponse = jsonResponse.Value<JObject>("data").Value<JObject>("ms.vss-work-web.backlogs-hub-backlog-data-provider").Value<JObject>("backlogPayload").Value<JObject>("queryResults");
        var sourceIds = jsonResponse.Value<JArray>("sourceIds");
        var targetIds = jsonResponse.Value<JArray>("targetIds");
        var azureItems = jsonResponse.Value<JObject>("payload").Value<JArray>("rows");
        foreach (var azureItem in azureItems)
        {
          var item = new DataItem();
          item.CreateThis = false;
          item.UpdateThis = true;
          string itemName = azureItem[1].ToString();
          if(itemName.Contains(' '))
          {
          item.Id = itemName.Substring(0, itemName.IndexOf(" "));
          item.Name = itemName.Substring(itemName.IndexOf(" ") + 1);
          item.Type = azureItem[0].ToString();
          item.AzureId = Convert.ToInt32(azureItem[7]);
          item.Revision = Convert.ToInt32(azureItem[9]);
          item.ParentId = (item.Type == "Epic") ? "" : item.Id.Substring(0, item.Id.LastIndexOf("."));
          OnlineBacklog.Add(item);
          }
        }
        return true;
      }
      catch (Exception ex) { MessageBox.Show(ex.Message, Resx.MessageError, MessageBoxButton.OK, MessageBoxImage.Error); return false; }
    }
    public async Task<bool> CreateAndUpdateDataItems(List<DataItem> items, List<DataItem> parents = null)
    {
      try
      {
        ErrorItems.Clear();
        string body = "";
        //Post Api Url bereitmachen
        string apiUrl = Url.Substring(0, Url.LastIndexOf("/_backlogs/")) + "/_api/_wit/updateWorkItems?__v=5";
        foreach (DataItem item in items)
        {
          if (!item.CreateThis && !item.UpdateThis) continue;
          body = (item.CreateThis) ? PrepareBodyForCreation(item, parents) : PrepareBodyForUpdate(item);
          var result = await BasePostRequestAsync(apiUrl, body);
          var response = await result.Content.ReadAsStringAsync();
          var jsonResponse = JObject.Parse(response).Value<JArray>("__wrappedArray")[0];
          if (jsonResponse.Value<string>("state").ToLower() == "error")
          {
            string errortext = (Resx.ApiConnectorErrorOnItem+" " + item.Id + " " + item.Name + "\n" + jsonResponse.Value<JObject>("error").Value<string>("message"));
            var Result = MessageBox.Show(errortext + "\n\n" + Resx.ApiConnectorErrorOnItemTryAgain, Resx.MessageAlert, MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (Result == MessageBoxResult.Yes)
            {
              item.AzureEmployee = "";
              body = (item.CreateThis) ? PrepareBodyForCreation(item, parents) : PrepareBodyForUpdate(item);
              result = await BasePostRequestAsync(apiUrl, body);
              response = await result.Content.ReadAsStringAsync();
              jsonResponse = JObject.Parse(response).Value<JArray>("__wrappedArray")[0];

              if (jsonResponse.Value<string>("state").ToLower() == "error")
              {
                errortext = (Resx.ApiConnectorErrorOnItem+" " + item.Id + " " + item.Name + "\n" + jsonResponse.Value<JObject>("error").Value<string>("message"));
                Result = MessageBox.Show(errortext + "\n\n" + Resx.ApiConnectorErrorOnItemContinue , Resx.MessageError, MessageBoxButton.YesNo, MessageBoxImage.Error);
                ErrorItems.Add(item.Id + " " + item.Name + "\n"+Resx.ApiConnectorErrorOnItemNotCreated);
                if (Result == MessageBoxResult.Yes) continue;
                else return false;
              }
              ErrorItems.Add(item.Id + " " + item.Name + "\n"+Resx.ApiConnectorErrorOnItemCreatedWOName);
            }
            else return false;
          }
          item.AzureId = jsonResponse.Value<int>("id");
          item.AzureDate = jsonResponse.Value<JObject>("fields").Value<string>("-5");
        }
        return true;
      }
      catch (Exception ex) { MessageBox.Show(ex.Message, Resx.MessageError, MessageBoxButton.OK, MessageBoxImage.Error); return false; }
    }
    public string PrepareBodyForCreation(DataItem item, List<DataItem> parents)
    {
      string body = createBody;
      body = body.Replace("itemType", item.Type);
      body = body.Replace("itemName", item.Id + " " + item.Name);
      body = body.Replace("employeeValue", item.AzureEmployee == "" ? "" : @"\""24\"":\""" + item.AzureEmployee);
      if (parents != null)
      {
        body = body.Replace("itemLink", addedLink);
        var parent = parents.Where(x => x.Id == item.ParentId).First();
        body = body.Replace("parentId", parent.AzureId.ToString());
        body = body.Replace("azDate", parent.AzureDate);
      }
      else body = body.Replace("itemLink", "");
      if (item.Type == "Task") body = body.Replace("itemValue", "");
      else body = body.Replace("itemValue", itemValue);
      return body;
    }
    public string PrepareBodyForUpdate(DataItem item)
    {
      return updateBody.Replace("itemName", item.Id + " " + item.Name).Replace("azureId", item.AzureId.ToString()).Replace("scopeValue", scopeValue).Replace("revision", item.Revision.ToString());
    }
    public async Task<HttpResponseMessage> BasePostRequestAsync(string apiUrl, string jsonRequestBody)
    {
      HttpContent contentbody = new StringContent(jsonRequestBody, Encoding.UTF8, "application/json");
      HttpResponseMessage response = await client.PostAsync(apiUrl, contentbody);
      return response;
    }
    public async Task<HttpResponseMessage> BaseGetRequestAsync(string apiUrl)
    {
      HttpResponseMessage response = await client.GetAsync(apiUrl);
      return response;
    }
    #endregion Methods
  }
}
