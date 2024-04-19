using AxcToAzure.Model;
using AxcToAzure.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AxcToAzure.ViewModel
{
  public class ItemSelectViewModel : NotifyObject
  {
    #region Properties
    public bool ItemSelectViewVisible
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public ObservableCollection<DataItem> DataItems
    {
      get { return Get<ObservableCollection<DataItem>>(); }
      set { Set(value); }
    }
    public ObservableCollection<DataItem> Nodes
    {
      get { return Get<ObservableCollection<DataItem>>(); }
      set { Set(value); }
    }

    public EventHandler<bool> Working;
    public event EventHandler<int> ChangeStep;
    #endregion
    #region Constructor
    public ItemSelectViewModel()
    {
      ItemSelectViewVisible = false;
      DataItems = new ObservableCollection<DataItem>();
      Nodes = new ObservableCollection<DataItem>();
      CreateCommands();
    }

    #endregion
    #region Methods
    public void BuildTree()
    {
    // Angezeigter Tree startet bei Epics
      Nodes = new ObservableCollection<DataItem>();
      foreach (var item in DataItems)
      {
        if (item.Type == "Epic") Nodes.Add(item);
      }

    }
    /// <summary>
    /// Wählt Parent aus, wenn Child gewählt wurde
    /// </summary>
    /// <param name="item"></param>
    private void SetParentsCreateStatus(DataItem item)
    {
      if (item.Type != "Epic")
      {
        var Parent = DataItems.Where(x => x.Id == item.ParentId).First();
        //Wenn Parent abgewählt war und Child ausgewählt => Wähle Parent aus
        if (!Parent.CreateThis && item.CreateThis)
        {
          Parent.CreateThis = true;
          SetParentsCreateStatus(Parent);
        }
      }
    }
    /// <summary>
    /// Wählt Kinder ab, wenn Parent abgewählt wurde
    /// </summary>
    /// <param name="item"></param>
    private void SetChildrensCreateStatus(DataItem item)
    {
      if (item.Children.Count > 0) //Gibt es Kinder?
      {
        if (!item.CreateThis) // Wurde Parent abgewählt?
        {

          foreach (var child in item.Children)
          {
            if (child.CreateThis) // Gibt es Kinder die noch angewählt sind?
            {

              child.CreateThis = false;
              SetChildrensCreateStatus(child);
            }
          }
        }
      }
    }
    #endregion
    #region Commands
    public ICommand CheckboxClickedCommand { get; private set; }
    public ICommand ContinueCommand { get; private set; }
    public ICommand BackCommand { get; private set; }
    public void CreateCommands()
    {
      CheckboxClickedCommand = new RelayCommand<string>(CheckboxClicked);
      ContinueCommand = new RelayCommand(Continue);
      BackCommand = new RelayCommand(Back);
    }
    /// <summary>
    /// Funktion um die Items an- und abzuwählen
    /// </summary>
    private void CheckboxClicked(string id)
    {
      var item = DataItems.Where(x => x.Id == id).First();
      SetParentsCreateStatus(item);
      SetChildrensCreateStatus(item);
    }
    private void Continue()
    {
      ChangeStep(this, 3);
    }
    private void Back()
    {
      ChangeStep(this, 1);
    }
    #endregion
  }
}
