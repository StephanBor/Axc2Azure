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

namespace xls2aturenet6.Model
{
  public class APIConnector
  {

    #region Properties
    private HttpClientHandler handler = new();
    private HttpClient client = new();
    public string Url { get; set; }
    public bool Initialized { get; set; }
    public object JsonConvert { get; private set; }

    public string epicBody = @"{
    ""updatePackage"": ""[{\""id\"":0,\""rev\"":0,\""projectId\"":\""\"",\""isDirty\"":true,\""tempId\"":-4,\""fields\"":{\""1\"":\""itemName\"",\""2\"":\""New\"",\""22\"":\""New\"",employeeValue\""25\"":\""Epic\"",\""10007\"":{\""type\"":1},\""10015\"":2,\""10018\"":\""Business\"",\""-2\"":apiTeamId,\""-104\"":apiProjectId}}]""
    }";

    public string featureBody = @"{
""updatePackage"":""[{\""id\"":0,\""rev\"":0,\""projectId\"":\""\"",\""isDirty\"":true,\""tempId\"":-4,\""fields\"":{\""1\"":\""itemName\"",\""2\"":\""New\"",\""22\"":\""New\"",employeeValue\""25\"":\""Feature\"",\""10007\"":{\""type\"":1},\""10015\"":2,\""10018\"":\""Business\"",\""-2\"":apiTeamId,\""-104\"":apiProjectId},\""links\"":{\""addedLinks\"":[{\""ID\"":parentId,\""LinkType\"":-2,\""Comment\"":\""\"",\""FldID\"":37,\""Changed Date\"":\""\\/azDate\\/\"",\""Revised Date\"":\""\\/azDate\\/\"",\""isAddedBySystem\"":true}]}}]""
}";

    public string storyBody = @"{
""updatePackage"":""[{\""id\"":0,\""rev\"":0,\""projectId\"":\""\"",\""isDirty\"":true,\""tempId\"":-1,\""fields\"":{\""1\"":\""itemName\"",\""2\"":\""New\"",\""22\"":\""New\"",employeeValue\""25\"":\""User Story\"",\""10007\"":{\""type\"":1},\""10015\"":2,\""10018\"":\""Business\"",\""-2\"":apiTeamId,\""-104\"":apiProjectId},\""links\"":{\""addedLinks\"":[{\""ID\"":parentId,\""LinkType\"":-2,\""Comment\"":\""\"",\""FldID\"":37,\""Changed Date\"":\""\\/azDate\\/\"",\""Revised Date\"":\""\\/azDate\\/\"",\""isAddedBySystem\"":true}]}}]""
}";

    public string taskBody = @"{
""updatePackage"":""[{\""id\"":0,\""rev\"":0,\""projectId\"":\""\"",\""isDirty\"":true,\""tempId\"":-2,\""fields\"":{\""1\"":\""itemName\"",\""2\"":\""New\"",\""22\"":\""New\"",employeeValue\""25\"":\""Task\"",\""10007\"":{\""type\"":1},\""10015\"":2,\""-2\"":apiTeamId,\""-104\"":apiProjectId},\""links\"":{\""addedLinks\"":[{\""ID\"":parentId,\""LinkType\"":-2,\""Comment\"":\""\"",\""FldID\"":37,\""Changed Date\"":\""\\/azDate\\/\"",\""Revised Date\"":\""\\/azDate\\/\"",\""isAddedBySystem\"":true}]}}]""
}";
    //public Dictionary<int, WorkProject> WorkProjects = new();
    #endregion Properties

    #region Construktor
    public APIConnector(string username, SecureString password, string url, string proxyAdress = "")
    {
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
        var teamBacklogName = Url.Substring(startindex, Url.LastIndexOf("/Epics") - startindex);
        var apiUrl = Url.Substring(0, Url.LastIndexOf("/_backlogs/")) + "?__rt=fps";
        //Finde scopeValue des Projekts für ApiUrl
        var result = await BaseGetRequestAsync(apiUrl);
        var response = await result.Content.ReadAsStringAsync();
        if (!result.IsSuccessStatusCode) { throw new Exception(result.ReasonPhrase + "\n" + response); }
        var jsonResponse = JObject.Parse(response);
        var scopeValue = jsonResponse["fps"].Value<JObject>("dataProviders").Value<string>("scopeValue");
        apiUrl = apiUrl.Substring(0, apiUrl.LastIndexOf("/")) + "/" + scopeValue + "/_api/_wit/nodes?__v=5";
        //Finde Projekt und Team id
        result = await BaseGetRequestAsync(apiUrl);
        response = await result.Content.ReadAsStringAsync();
        if (!result.IsSuccessStatusCode) { throw new Exception(result.ReasonPhrase + "\n" + response); }
        jsonResponse = JObject.Parse(response);
        string teamId = jsonResponse.Value<JArray>("children")[0].Value<JArray>("children").Where(x => x.Value<string>("name") == teamBacklogName).First().Value<string>("id");
        string projectId = jsonResponse.Value<string>("id");
        // Post Bodies vorbereiten
        epicBody = epicBody.Replace("apiTeamId", teamId).Replace("apiProjectId", projectId);
        featureBody = featureBody.Replace("apiTeamId", teamId).Replace("apiProjectId", projectId);
        storyBody = storyBody.Replace("apiTeamId", teamId).Replace("apiProjectId", projectId);
        taskBody = taskBody.Replace("apiTeamId", teamId).Replace("apiProjectId", projectId);
        Initialized = true;
        return true;
      }
      catch (Exception ex) { MessageBox.Show(ex.Message); return false; }
    }
    public async Task<bool> WorkData(List<DataItem> items, List<DataItem> parents = null)
    {
      try
      {
        string body = "";
        //Post Api Url bereitmachen
        string apiUrl = Url.Substring(0, Url.LastIndexOf("/_backlogs/")) + "/_api/_wit/updateWorkItems?__v=5";
        foreach (DataItem item in items)
        {
          body = PrepareBody(item, parents);
          var result = await BasePostRequestAsync(apiUrl, body);
          var response = await result.Content.ReadAsStringAsync();
          var jsonResponse = JObject.Parse(response).Value<JArray>("__wrappedArray")[0];
          if (jsonResponse.Value<string>("state").ToLower() == "error")
          {
            string errortext = ("Error on item: " + item.Id + " " + item.Name + "\n" + jsonResponse.Value<JObject>("error").Value<string>("message"));
            var Result = MessageBox.Show(errortext, "Would you like to try again (without an employee set)?", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (Result == MessageBoxResult.Yes)
            {
              item.Employee = "";
              body = PrepareBody(item, parents);
              result = await BasePostRequestAsync(apiUrl, body);
              response = await result.Content.ReadAsStringAsync();
              jsonResponse = JObject.Parse(response).Value<JArray>("__wrappedArray")[0];
              if (jsonResponse.Value<string>("state").ToLower() == "error")
              {
                errortext = ("Error on item: " + item.Id + " " + item.Name + "\n" + jsonResponse.Value<JObject>("error").Value<string>("message"));
              }
            }
            Result = MessageBox.Show(errortext, "Would you like to continue with the next item?", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (Result == MessageBoxResult.Yes) continue;
            else return false;
          }
          item.AzureId = jsonResponse.Value<int>("id");
          item.AzureDate = jsonResponse.Value<JObject>("fields").Value<string>("-5");
        }
        return true;
      }
      catch (Exception ex) { MessageBox.Show(ex.Message); return false; }
    }
    public string PrepareBody (DataItem item, List<DataItem> parents)
    {
      string body = "";
      switch (item.Type)
      {
        case "epic":
          body = epicBody;

          break;
        case "feature":
          body = featureBody;

          break;
        case "story":
          body = storyBody;

          break;
        case "task":
          body = taskBody;

          break;
        default: throw new Exception("Problem with " + item.Type + " " + item.Name);
      }
      body = body.Replace("itemName", item.Name);
      body = body.Replace("employeeValue", item.Employee);
      if (parents != null)
      {
        foreach (var parent in parents)
        {
          if (parent.Id == item.ParentId)
          {
            body = body.Replace("parentId", parent.AzureId.ToString());
            body = body.Replace("azDate", parent.AzureDate);
            break;
          }
        }
      }
      return body;
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
