using AxcToAzure.Model;
using AxcToAzure.Utilities;
using AxcToAzure.View;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AxcToAzure.ViewModel
{
  public class ExcelViewModel : NotifyObject
  {
    #region Properties
    public bool ExcelViewVisible
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public string FilePath
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public ObservableCollection<string> WorksheetNames
    {
      get { return Get<ObservableCollection<string>>(); }
      set { Set(value); }
    }
    public ObservableCollection<DataItem> DataItems
    {
      get { return Get<ObservableCollection<DataItem>>(); }
      set { Set(value); }
    }
    public string WorksheetName
    {
      get { return Get<string>(); }
      set { Set(value); SheetSelected = (value != "" && value != null); CanContinue = false; }
    }
    public string NumberColumn
    {
      get { return Get<string>(); }
      set { Set(value.ToUpper().Trim()); CheckColumnsValue();  }
    }
    public string DescriptionColumn
    {
      get { return Get<string>(); }
      set { Set(value.ToUpper().Trim()); CheckColumnsValue(); }
    }
    public string EmployeeColumn
    {
      get { return Get<string>(); }
      set { Set(value.ToUpper().Trim()); CheckColumnsValue(); }
    }
    public int DefaultEmployee
    {
      get { return Get<int>(); }
      set {  Set(value); }

    }
    public bool FileLoaded
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool SheetSelected
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool CanContinue
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public string ItemWorkedOn
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public double BarProgress
    {
      get { return Get<double>(); }
      set { Set(value); }
    }
    public bool ColumnsSet
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    private Excel.Workbook workBook;
    public EventHandler<bool> Working;
    public event EventHandler<int> ChangeStep;
    public bool FileInReading
    {
      get { return Get<bool>(); }
      set { Set(value); }

    }

    #endregion
    #region Constructor
    public ExcelViewModel()
    {

      WorksheetNames = new ObservableCollection<string>();
      ExcelViewVisible = true;
      FileLoaded = false;
      SheetSelected = false;
      ColumnsSet = false; 
      FileInReading = false;
      CanContinue = false;
      DescriptionColumn = "";
      NumberColumn = "";
      EmployeeColumn = "";
      DefaultEmployee = 1;
      CreateCommands();
    }
    #endregion
    #region Methods
    /// <summary>
    /// Überträgt die Namen der Worksheets in die Liste für die Combobox
    /// </summary>
    public void OpenFile()
    {
      Working.Invoke(this, true);
      WorksheetName = "";
      WorksheetNames =  new ObservableCollection<string>();
      SheetSelected = false;
      CanContinue = false;

      Excel.Application excel = new Excel.Application();
      workBook = excel.Workbooks.Open(FilePath);
      for (int i = 1; i <= workBook.Worksheets.Count; i++)
      {

      WorksheetNames.Add( ((Worksheet)workBook.Worksheets[i]).Name);
      }
     
      FileLoaded = true;
      Working.Invoke(this, false) ;
    }
    /// <summary>
    /// Prüft ob beide Columns angegeben sind und ob deren Werte eine brauchbare Excel Spalte haben (von A - ZZ)
    /// </summary>
    public void CheckColumnsValue()
    {
      CanContinue = false;
      Regex rx = new Regex(@"\A[A-Z]{1,2}\Z"); 
      if (NumberColumn == null || DescriptionColumn == null) return;
      if ( EmployeeColumn == null || EmployeeColumn =="")
      {
        bool IsMatch = (DescriptionColumn == NumberColumn);
        ColumnsSet = (rx.IsMatch(DescriptionColumn) && rx.IsMatch(NumberColumn) && !IsMatch);
      }
      else
      {
      bool IsMatch = (DescriptionColumn == NumberColumn) || (DescriptionColumn == EmployeeColumn) || (EmployeeColumn == NumberColumn);
          ColumnsSet = (rx.IsMatch(DescriptionColumn) && rx.IsMatch(NumberColumn) && rx.IsMatch(EmployeeColumn) && !IsMatch);
      }
    }
    /// <summary>
    /// Verändert den Mousecursor bei längerem Laden
    /// </summary>
    public void sortData()
    {
      FileInReading = true;
      CanContinue = false;
      DataItems = new ObservableCollection<DataItem>();
      bool useEmployees = (EmployeeColumn != null && EmployeeColumn != "");
      try
      {
        // Holt sich das richtige Arbeitsblatt
        Excel.Worksheet ws = (Worksheet)workBook.Worksheets[WorksheetName];  
        // Holt sich die Anzahl der Zeilen
        int rows = ws.UsedRange.Rows.Count;
        // Blueprint für die Benennung anlegen
        Regex epicReg = new Regex(@"^\d+\Z");
        Regex featureReg = new Regex(@"^\d+\.\d+\Z");
        Regex storyReg = new Regex(@"^\d+\.\d+\.\d+\Z");
        Regex taskReg = new Regex(@"^\d+\.\d+\.\d+\.\d+\Z");
        Regex employeeReg = new Regex(@"^[A-Z][a-z]+\,[A-Z][a-z]+(\;[A-Z][a-z]+\,[A-Z][a-z]+)*\Z");
        ItemWorkedOn = "";
        BarProgress = 0;
        //Sortierschleife
        for (int i = 1; i <= rows; i++)
        {
          string objectId = ws.Range[(NumberColumn+i).ToString()].Text.ToString();
          double progress = i *100/ (rows+1);
          if (i != rows)
          {
            ItemWorkedOn = "Reading: "+objectId + " (" + (progress) + " %)";
            BarProgress = progress;
          }
          else
          {
            BarProgress = 99;
            ItemWorkedOn = "Setting Children for Data...";

          }
          string objectName = ws.Range[(DescriptionColumn + i).ToString()].Text.ToString();
          string objectEmployee = "";
          if (useEmployees)
          {
            objectEmployee = ws.Range[(EmployeeColumn + i).ToString()].Text.ToString();
            bool IsMatch = employeeReg.IsMatch(Regex.Replace(objectEmployee, @"\s+", string.Empty));
            if (!IsMatch) objectEmployee = ""; // Replace whitespace characters
          }
          if (epicReg.IsMatch(objectId))
          {
            DataItems.Add(CreateDataItem(objectId, objectName, objectEmployee, "Epic"));
          }
          if (featureReg.IsMatch(objectId))
          {
            DataItems.Add(CreateDataItem(objectId, objectName, objectEmployee, "Feature"));
          }
          if (storyReg.IsMatch(objectId))
          {
            DataItems.Add(CreateDataItem(objectId, objectName, objectEmployee, "User Story"));
          }
          if (taskReg.IsMatch(objectId))
          {
            DataItems.Add(CreateDataItem(objectId, objectName, objectEmployee, "Task"));
          }
        }
        SetItemChildren();
        BarProgress = 100;
        ItemWorkedOn = "Finished";
        
      }
      catch (Exception ex) {
        ItemWorkedOn = "Error";
        BarProgress = 0; 
        MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error); 
      }
      FileInReading = false;
      CanContinue = (ItemWorkedOn != "Error");
    }
    
    public DataItem CreateDataItem(string testId, string name, string employee, string type)
    {
      DataItem dataItem = new DataItem();
      dataItem.Id = testId;
      dataItem.Name = name.Replace("\"", "\'");
      dataItem.Type = type;
      dataItem.ParentId = "";
      dataItem.Employee = employee;
      if (type != "Epic")
      {
        dataItem.ParentId = testId.Substring(0, testId.LastIndexOf("."));
      }
      //if (debug) Console.WriteLine(dataItem.Id + " " + dataItem.Name + " " + dataItem.ParentId);
      return dataItem;
    }
    public void SetItemChildren()
    {
      
      foreach (var item in DataItems)
      {
        if(item.Type != "Task")
        {
          
          foreach (var child in DataItems)
          {
            if (child.Type != "Epic" && item.Id == child.ParentId)
            {
              item.Children.Add(child);
            }
          }
        }
      }
    }
    #endregion
    #region Commands
    public ICommand SelectFileCommand { get; private set; }
    public ICommand OpenInstructionCommand { get; private set; }
    public ICommand ReadFileCommand { get; private set; }
    public ICommand ChangeDefaultEmployeeCommand { get; private set; }
    public ICommand ContinueCommand { get; private set; }
    public ICommand ExitCommand { get; private set; }
    public void CreateCommands()
    {
      SelectFileCommand = new RelayCommand(SelectFile);
      OpenInstructionCommand = new RelayCommand(OpenInstruction);
      ReadFileCommand = new RelayCommand(ReadFile);
      ChangeDefaultEmployeeCommand = new RelayCommand<string>(ChangeDefaultEmployee);
      ContinueCommand = new RelayCommand(Continue);
      ExitCommand = new RelayCommand(Exit);
    }
    /// <summary>
    /// Funktion um die Excel Datei zu finden
    /// </summary>
    private void SelectFile()
    {
      if (FileInReading) return;
      OpenFileDialog selectFileDialog = new OpenFileDialog();
      selectFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
      if (selectFileDialog.ShowDialog() == true)
      {
        FilePath = selectFileDialog.FileName;
        OpenFile();
      }
    }
    private void OpenInstruction()
    {
      InstructionWindow instructionWindow = new InstructionWindow();
      instructionWindow.Show();
    }
    private void ReadFile()
    {
      if (!FileInReading)
      {
        Thread t = new Thread(sortData);
        t.Start();
      }
    }
    private void ChangeDefaultEmployee (string inc)
    {
      int x = Convert.ToInt32(inc);
      if (FileInReading || (DefaultEmployee + x)< 1) return;
      DefaultEmployee += x;
    }
    private void Continue()
    {
      if (!FileInReading)
      {
        ChangeStep.Invoke(this, 2);
      }
    }
      private void Exit()
    {
      Environment.Exit(0);
    }
    #endregion

  }
}
