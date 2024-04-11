using AxcToAzure.Model;
using AxcToAzure.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using xls2aturenet6.Model;

namespace AxcToAzure.ViewModel
{
  public class ProgressViewModel : NotifyObject
  {
    #region Properties
    public bool ProgressViewVisible
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool CompareBacklogs
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool DataWorking
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool FinishedSuccessfully
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
    public APIConnector ApiConnector
    {
      get { return Get<APIConnector>(); }
      set { Set(value); }
    }
    public EventHandler<bool> Working;
    public event EventHandler<int> ChangeStep;
    private List<DataItem> epics;
    private List<DataItem> features;
    private List<DataItem> stories;
    private List<DataItem> tasks;
    #endregion
    #region Constructor
    public ProgressViewModel()
    {
      ProgressViewVisible = false;
      FinishedSuccessfully = false;
       
      DataWorking = false;
      BarProgress = 0;
      DataItems = new ObservableCollection<DataItem>();
      CreateCommands();
    }

    #endregion
    #region Methods
    public void SortData()
    {
      epics = new List<DataItem>();
      features = new List<DataItem>();
      stories = new List<DataItem>();
      tasks = new List<DataItem>();
      foreach(var item in DataItems)
      {
        if (!item.CreateThis) continue;
        switch (item.Type)
        {
          case "Epic": 
            epics.Add(item); break;
          case "Feature":
            features.Add(item); break;
          case "User Story":
            stories.Add(item); break;
          case "Task":
            tasks.Add(item); break;
        }
      }
    }
    async Task<bool> WorkData()
    {
      Log = "";
      string currentItemClass = "";
      DataWorking = true;
      List<DataItem> dataItems = new List<DataItem>();
      List<DataItem> parents = new List<DataItem>();
      BarProgress = 0;
      if (!ApiConnector.Initialized)
      {
        Log += "Establishing Connection to Backlog\n";
        if (!await ApiConnector.InitializeConnection()) 
        { 
          Log += "Error. Check your URL, Credentials and Internet Connection\n";
          DataWorking = false;
          return false; 
        }
        Log += "Connection Successfully established\n";
      }
      for (int i = 0; i < 4; i++)
      {
        switch (i)
        {
          case 0:
            currentItemClass = "Epics";
            dataItems = epics;
            parents = null;
            break;
          case 1:
            currentItemClass = "Features";
            dataItems = features;
            parents = epics;
            break;
          case 2:
            currentItemClass = "Stories";
            dataItems = stories;
            parents = features;
            break;
          case 3:
            currentItemClass = "Tasks";
            dataItems = tasks;
            parents = stories;
            break;
          default: break;

        }
        Log+="Start with creating " + currentItemClass+"\n";
        if (!await ApiConnector.WorkData(dataItems, parents))
        {
          BarProgress = 0;
          Log += "An Error occured. Please check your Internet Connection.\n";
          DataWorking = false;
          FinishedSuccessfully=false; 
          return false;

        }
        BarProgress += 25;
        Log += currentItemClass + " created successfully.\n";
      }
      DataWorking = false;
      FinishedSuccessfully = true;
      return true;

    }

    #endregion
    #region Commands
    public ICommand StartCommand { get; private set; }
    public ICommand ExitCommand { get; private set; }
    public ICommand BackCommand { get; private set; }
    public void CreateCommands()
    {
      StartCommand = new RelayCommand(Start);
      ExitCommand = new RelayCommand(Exit);
      BackCommand = new RelayCommand(Back);
    }
    private  void Start()
    {
      if (DataWorking) return;
      Working.Invoke(this, true);
      new Thread(()=>WorkData().Wait()).Start();
      Working.Invoke(this, false);
    }
   
    private void Exit()
    {
      Environment.Exit(0);
    }
    private void Back()
    {
      if (DataWorking) return;
      if (CompareBacklogs) ChangeStep(this, 5);
      else ChangeStep(this, 4);
    }
    #endregion
  }
}