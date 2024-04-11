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
    public void FilterCreatableItems(ObservableCollection<DataItem> items, int defaultEmployee)
    {
      DataItems = new ObservableCollection<DataItem>();
      foreach (var item in items)
      {
        if (item.CreateThis) DataItems.Add(item);
      }
      SetItemEmployees(defaultEmployee);
    }
    public void SetItemEmployees(int defaultEmployee)
    {
      List<DataItem> newTasks = new List<DataItem>();
      foreach (var item in DataItems)
      {
        if (item.Employee == "") continue;
        string[] employees = item.Employee.Split(";");
        if (item.Type == "Task")
        { //Copy the task
          for (int i = 0; i < employees.Length; i++)
          {
            string[] names = employees[i].Split(",");
            string firstname = names[1].Trim();
            string lastname = names[0].Trim();
            string axcName = @"\""24\"":\""" + firstname + " " + lastname + @"<PROLEIT-AG\\\\" + firstname + "_" + lastname + @">\"",";
            if (i == 0) //Setze Employee für schon bestehenden task, danach kreire neue
            {
              item.AzureEmployee = axcName;
          
              continue;
            }
            DataItem task = new DataItem();
            task.Id = item.Id;
            task.Name = item.Name;
            task.ParentId = item.ParentId;
            task.Type = item.Type;
            task.Employee = axcName;
            task.AzureEmployee = axcName;
            newTasks.Add(task);
          }
        }
        else
        {
          //Wenn Defaultemployees zu hoch eingestellt ist, nimm letzten Employee
          int emplid = (defaultEmployee > employees.Length) ? employees.Length - 1 : defaultEmployee - 1 ;
          string[] names = employees[emplid].Split(","); 
          string firstname = names[1].Trim();
          string lastname = names[0].Trim();
          string axcName = @"\""24\"":\"""+firstname + " " + lastname + @"<PROLEIT-AG\\\\" + firstname + "_" + lastname + @">\"",";
          item.Employee = axcName;
          item.AzureEmployee = axcName;
        }
      }
      //Nun füge neu angelegte Tasks wieder ein
      foreach (var task in newTasks)
      {
        DataItems.Add(task);
        DataItems.Where(x => x.Id == task.ParentId).First().Children.Add(task);
      }
    }
    public void SelectUserStories()
    {
      Stories = new ObservableCollection<DataItem>();
      foreach (var item in DataItems)
      {
        if (item.Type == "User Story") Stories.Add(item);
      }
    }
    private void AddDefaultTasksToStory(string[] newTasks, DataItem story)
    {
      int newChildIndex = GetLastChildIndexOfStory(story.Children) + 1;
      foreach (var task in newTasks)
      {
        if (task != null && task.Trim() != "")
        {
          DataItem newTask = new DataItem();
          newTask.Id = story.Id + "." + newChildIndex;
          newTask.Name = task.Trim();
          newTask.Type = "Task";
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
      foreach (var child in Children)
      {
        int compareIndex = Convert.ToInt32(child.Id.Substring(child.Id.LastIndexOf(".") + 1));
        if (compareIndex > lastIndex)
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
      catch (Exception ex) { MessageBox.Show(ex.ToString()); }
      workBook.Close();
      Working.Invoke(this, false);
    }
    private void Continue()
    {
      Working.Invoke(this, true);
      //ChangeStep(this, 4);
      foreach (var story in Stories)
      {
        if (story.DefaultTask != "" && story.DefaultTask != null)
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
