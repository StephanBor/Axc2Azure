using AxcToAzure.Model;
using AxcToAzure.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace AxcToAzure.ViewModel
{
  public class DefaultTaskViewModel : NotifyObject
  {
    #region Properties
    public bool DefaultTaskViewVisible
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public ObservableCollection<DataItem> DataItems
    {
      get { return Get<ObservableCollection<DataItem>>(); }
      set { Set(value); }
    }
    public ObservableCollection<DataItem> Stories
    {
      get { return Get<ObservableCollection<DataItem>>(); }
      set { Set(value); }
    }
    public Dictionary<string, string> DefaultTaskList
    {
      get { return Get<Dictionary<string, string>>(); }
      set { Set(value); }
    }
    private string pathToDefaultTaskExcel;
    public EventHandler<bool> Working;
    public event EventHandler<int> ChangeStep;
    #endregion
    #region Constructor
    public DefaultTaskViewModel()
    {
      DefaultTaskViewVisible = false;
      pathToDefaultTaskExcel = System.AppDomain.CurrentDomain.BaseDirectory + @"\Utilities\Default_Tasks.xlsx";
      DataItems = new ObservableCollection<DataItem>();
      Stories = new ObservableCollection<DataItem>();
      CreateCommands();
    }
    #endregion
    #region Methods
    public void FilterCreatableItems(ObservableCollection<DataItem> items)
    {
      DataItems = new ObservableCollection<DataItem>();
      foreach(var item in items)
      {
        if(item.CreateThis) DataItems.Add(item);
      }
    }
    public void SelectUserStories()
    {
      Stories = new ObservableCollection<DataItem>();
      foreach (var item in DataItems)
      {
        if (item.Type == "story") Stories.Add(item);
      }
    }
    private void AddDefaultTasksToStory(string[] newTasks, DataItem story)
    {
      int newChildIndex= GetLastChildIndexOfStory(story.Children)+1;
      foreach (var task in newTasks)
      {
        if (task != null && task.Trim() != "")
        {
          DataItem newTask = new DataItem();
          newTask.Id = story.Id + "." + newChildIndex;
          newTask.Name = task.Trim();
          newTask.Type = "task";
          newTask.ParentId = story.Id;
          story.Children.Add(newTask);
          DataItems.Add(newTask);
          newChildIndex += 1;
        }
      }
    }
    private int GetLastChildIndexOfStory(ObservableCollection<DataItem> Children)
    {
      int lastIndex = 0;
      foreach( var child in Children)
      {
        int compareIndex = Convert.ToInt32(child.Id.Substring(child.Id.LastIndexOf(".")+1));
        if(compareIndex > lastIndex)
        {
          lastIndex = compareIndex;
        }
      }
      return lastIndex;
    }
    #endregion
    #region Commands
    public ICommand OpenDefaultTaskListCommand { get; private set; }
    public ICommand RefreshDefaultTaskListCommand { get; private set; }
    public ICommand ContinueCommand { get; private set; }
    public ICommand BackCommand { get; private set; }
    public void CreateCommands()
    {
      OpenDefaultTaskListCommand = new RelayCommand(OpenDefaultTaskList);
      RefreshDefaultTaskListCommand = new RelayCommand(RefreshDefaultTaskList);
      ContinueCommand = new RelayCommand(Continue);
      BackCommand = new RelayCommand(Back);
    }
    private void OpenDefaultTaskList()
    {
      Working.Invoke(this, true);
      ProcessStartInfo ps = new ProcessStartInfo();
      ps.FileName = "excel"; // "EXCEL.EXE" also works
      ps.Arguments = pathToDefaultTaskExcel;
      ps.UseShellExecute = true;
      Process.Start(ps);
      Working.Invoke(this, false);
    }
    public void RefreshDefaultTaskList()
    {
      Working.Invoke(this, true);
      DefaultTaskList = new Dictionary<string, string>();
      DefaultTaskList.Add("", "");
      Excel.Application excel = new Excel.Application();
      Excel.Workbook workBook = excel.Workbooks.Open(pathToDefaultTaskExcel);
      try
      {
        // Holt sich das richtige Arbeitsblatt
        Excel.Worksheet ws = (Worksheet)workBook.Worksheets[1];
        // Holt sich die Anzahl der Zeilen
        int rows = ws.UsedRange.Rows.Count;
        Regex taskReg = new Regex(@"^\d+\.\d+\.\d+\.\d+\Z");
        //Sortierschleife Start bei 2 wegen Header
        for (int i = 2; i <= rows; i++)
        {
          string key = ws.Range[("A" + i).ToString()].Text.ToString();
          string value = ws.Range[("B" + i).ToString()].Text.ToString();
          DefaultTaskList.Add(key, value);
        }
      }
      catch(Exception ex) { MessageBox.Show(ex.ToString()); }
      workBook.Close();
          Working.Invoke(this, false);
    }
    private void Continue()
    {
      Working.Invoke(this, true);
      //ChangeStep(this, 4);
      foreach (var story in Stories)
      {
        if(story.DefaultTask != "" && story.DefaultTask != null)
        {
          AddDefaultTasksToStory(story.DefaultTask.Split(";"), story);
        }
      }
      Working.Invoke(this, false);
      ChangeStep(this, 4);
    }
    private void Back()
    {
      Working.Invoke(this, true);
      Working.Invoke(this, false);
      ChangeStep(this, 2);
    }
    #endregion
  }

}
