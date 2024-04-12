using AxcToAzure.Model;
using AxcToAzure.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
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
    public ObservableCollection<DataItem> AzureDataItems
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
      AzureDataItems = new ObservableCollection<DataItem>();
      CreateCommands();
    }
    #endregion
    #region Methods
    async Task<bool> ReadBacklogData()
    {
      Log = "";
      BarProgress = 0;
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

      BarProgress = 100;
      Log = "Finished";
      BacklogInReading = false;
      CanContinue = true;
      return true;

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
      
      Working.Invoke(this, true);
      new Thread(() => ReadBacklogData().Wait()).Start();
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
