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
using Resx = AxcToAzure.Properties.Resources;


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
    /// <summary>
    /// Sortiert die Dataitems in 4 Listen, je nach Typ (Epic, Feature,...)
    /// </summary>
    public void SortData()
    {
      epics = new List<DataItem>();
      features = new List<DataItem>();
      stories = new List<DataItem>();
      tasks = new List<DataItem>();
      foreach(var item in DataItems)
      {
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
    /// <summary>
    /// Steuert den Ablauf um Items anzulegen
    /// </summary>
    /// <returns>true wenn erfolgreich</returns>
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
        //Initilisiere Verbindung zu DevOps
        Log += Resx.MessageEstablishConnection+"\n";
        if (!await ApiConnector.InitializeConnection())
        {
          Log += Resx.MessageLoginError + "\n";
          DataWorking = false;
          return false;
        }
        Log += Resx.MessageLoginSuccess + "\n";
      }
      //Setzt welcher Itemtyp bearbeitet werden soll und welcher Typ die Eltern sind
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
        Log+= Resx.ProgressViewModelStartCreating+" " + currentItemClass+"\n";
        //Übergebe zu Bearbeitende Daten and APIConnector
        if (!await ApiConnector.CreateAndUpdateDataItems(dataItems, parents))
        {
          BarProgress = 0;
          Log += Resx.ProgressViewModelErrorCreating+"\n";
          DataWorking = false;
          FinishedSuccessfully=false; 
          return false;

        }
        BarProgress += 25;
        //Zeige ob und welche Items nicht angelegt werden konnten
        if (ApiConnector.ErrorItems.Count() > 0) 
        {
          Log += currentItemClass + " "+Resx.ProgressViewModelPartiallyCreating +"\n";
          foreach (var item in ApiConnector.ErrorItems)
          {
            Log += item + "\n";
          }
        }
        else Log += currentItemClass + " "+ Resx.ProgressViewModelSuccessCreating + "\n";
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
      // neuer Thread um UI während der Bearbeitung upzudaten (Progressbar, Log)
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