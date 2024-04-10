using AxcToAzure.Model;
using AxcToAzure.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
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
    #endregion
    #region Constructor
    public BacklogsCompareViewModel()
    {
      BacklogsCompareViewVisible = false;
      CompareBacklogs = false;
      DataItems = new ObservableCollection<DataItem>();
      CreateCommands();
    }
    #endregion
    #region Methods

    #endregion
    #region Commands
    public ICommand ContinueCommand { get; private set; }
    public ICommand BackCommand { get; private set; }
    public void CreateCommands()
    {
      ContinueCommand = new RelayCommand(Continue);
      BackCommand = new RelayCommand(Back);
    }
    private void Continue()
    {
      CompareBacklogs = true;
      ChangeStep(this, 6);
    }
    private void Back()
    {
      CompareBacklogs = false;
     ChangeStep(this, 4);
    }
    #endregion
  }
}
