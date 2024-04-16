using AxcToAzure.Model;
using AxcToAzure.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using xls2aturenet6.Model;

namespace AxcToAzure.ViewModel
{
  public class BacklogsCompareViewModel : NotifyObject
  {
    #region Properties
    public bool BacklogsCompareViewVisible
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool CompareBacklogs
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool BacklogInReading
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool ShowProgress
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool CanContinue
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public string Log
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public double BarProgress
    {
      get { return Get<double>(); }
      set { Set(value); }
    }
    public ObservableCollection<DataItem> DataItems
    {
      get { return Get<ObservableCollection<DataItem>>(); }
      set { Set(value); }
    }
    public ObservableCollection<DataItem> ItemsToCompare
    {
      get { return Get<ObservableCollection<DataItem>>(); }
      set { Set(value); }
    }
    public APIConnector ApiConnector
    {
      get { return Get<APIConnector>(); }
      set { Set(value); }
    }
    public EventHandler<bool> Working;
    public event EventHandler<int> ChangeStep;
    
    #endregion
    #region Constructor
    public BacklogsCompareViewModel()
    {
      BacklogsCompareViewVisible = false;
      CompareBacklogs = false;
      CanContinue = false;
      BacklogInReading = false;
      ShowProgress = false;
      DataItems = new ObservableCollection<DataItem>();
      ItemsToCompare = new ObservableCollection<DataItem>();
      CreateCommands();
    }
    #endregion
    #region Methods
    async Task<bool> ReadBacklogData()
    {
      Log = "";
      BarProgress = 0;
      CanContinue = false;
      BacklogInReading = true;
      ShowProgress = true;
      if (!ApiConnector.Initialized)
      {
        Log = "Establishing Connection to Backlog";
        if (!await ApiConnector.InitializeConnection())
        {
          Log = "Error. Check your URL, Credentials and Internet Connection";
          BacklogInReading = false;
          return false;
        }
        Log = "Connection Successfully established";
      }
      BarProgress = 25;
      Log = "Trying to get existing Backlog";
      if ( !await ApiConnector.GetExistingBacklog())
      {
        Log = "Error while Reading the Backlog";
        BarProgress = 0;
        BacklogInReading = false;
        return false;
      }
      BarProgress = 75;
      Log = "Backlog successfully read. Comparing Data..";
      CompareItemsWithBacklog(ApiConnector.OnlineBacklog);
      BarProgress = 100;
      Log = "Finished";
      BacklogInReading = false;
      CanContinue = true;
      return true;

    }
    public void CompareItemsWithBacklog(List<DataItem> OnlineBacklog)
    {
      foreach (var item in DataItems)
      {
        if (!item.CreateThis && !item.UpdateThis) continue;
        if(item.Type != "Task") 
        {
          //Suche mögliche Partner
          var partner =OnlineBacklog.Where(x=> x.Id == item.Id).FirstOrDefault();
          //Kein partner gefunden => Item muss neu angelegt werden
          if(partner == null)  continue;
          //Items gleich => Mache nichts
            item.CreateThis = false;
          var oldEmployee = partner.AzureEmployee.Trim() == "" ? "Not set" : partner.AzureEmployee.Substring(0, partner.AzureEmployee.IndexOf("<"));
          var newEmployee = item.AzureEmployee.Trim() == "" ? "Not set" : item.AzureEmployee.Substring(0, item.AzureEmployee.IndexOf("<"));
          bool namesMatch = partner.Name == item.Name;
          bool employeesMatch = oldEmployee == newEmployee;
          if (namesMatch && employeesMatch)   continue;
          item.UpdateThis = true;
          item.AzureId = partner.AzureId;
          item.Revision = partner.Revision;
          if (!namesMatch) item.UpdateReason += $"Old Name: {partner.Name}\n";
          if (!employeesMatch) item.UpdateReason += $"old Employee: {oldEmployee}, new Employee: {newEmployee}";
          App.Current.Dispatcher.Invoke((Action)delegate // <--- HERE
          {
          ItemsToCompare.Add(item);
            
          });
        }
        else
        {

        }
      }
    }
    #endregion
    #region Commands
    public ICommand ContinueCommand { get; private set; }
    public ICommand ReadBacklogsCommand { get; private set; }
    public ICommand BackCommand { get; private set; }
    public void CreateCommands()
    {
      ContinueCommand = new RelayCommand(Continue);
      ReadBacklogsCommand = new RelayCommand(ReadBacklogs);
      BackCommand = new RelayCommand(Back);
    }
    private void Continue()
    {
      if (BacklogInReading) return;
      CompareBacklogs = true;
      ChangeStep(this, 6);
    }
    private void ReadBacklogs()
    {
      if (BacklogInReading) return;
      ItemsToCompare = new ObservableCollection<DataItem>();
      Working.Invoke(this, true);
      new Thread(() => ReadBacklogData().Wait()).Start();
      //ReadBacklogData().Wait();
      Working.Invoke(this, false);
    }
    private void Back()
    {
      if (BacklogInReading) return;
      CompareBacklogs = false;
     ChangeStep(this, 4);
    }
    #endregion
  }
}
