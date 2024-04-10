using AxcToAzure.Model;
using AxcToAzure.Utilities;
using AxcToAzure.View;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using xls2aturenet6.Model;

namespace AxcToAzure.ViewModel
{
  public class MainWindowViewModel : NotifyObject
  {
    #region Properties
    public ExcelViewModel ExcelViewModel
    {
      get { return Get<ExcelViewModel>(); }
      set { Set(value); }
    }
    public ItemSelectViewModel ItemSelectViewModel
    {
      get { return Get<ItemSelectViewModel>(); }
      set { Set(value); }
    }
    public DefaultTaskViewModel DefaultTaskViewModel
    {
      get { return Get<DefaultTaskViewModel>(); }
      set { Set(value); }
    }
    public LoginViewModel LoginViewModel
    {
      get { return Get<LoginViewModel>(); }
      set { Set(value); }
    }
    public BacklogsCompareViewModel BacklogsCompareViewModel
    {
      get { return Get<BacklogsCompareViewModel>(); }
      set { Set(value); }
    }
    public ProgressViewModel ProgressViewModel
    {
      get { return Get<ProgressViewModel>(); }
      set { Set(value); }
    }
    #endregion
    public MainWindowViewModel()
    {
      MainWindowViewModel mainWindow = this;
      ExcelViewModel = new ExcelViewModel();
      ItemSelectViewModel = new ItemSelectViewModel();
      DefaultTaskViewModel = new DefaultTaskViewModel();
      LoginViewModel = new LoginViewModel();
      BacklogsCompareViewModel = new BacklogsCompareViewModel();
      ProgressViewModel = new ProgressViewModel();
      ExcelViewModel.ChangeStep += ChangeStep;
      ExcelViewModel.Working += SetMouseLoading;
    }
    private void ChangeStep(object sender, int step)
    {
      switch (step)
      {
        case 1: //ExcelView
          //Alte Events abwählen
          ItemSelectViewModel.ChangeStep -= ChangeStep;
          ItemSelectViewModel.Working -= SetMouseLoading;
          //Neue Events hinzufügen
          ExcelViewModel.ChangeStep += ChangeStep;
          ExcelViewModel.Working += SetMouseLoading;
          //Alte Views abwählen und neue hinzufügen
          ItemSelectViewModel.ItemSelectViewVisible = false;
          ExcelViewModel.ExcelViewVisible = true;
          break;
        case 2: // ItemSelectView
          //Alte Events abwählen
          ExcelViewModel.ChangeStep -= ChangeStep;
          ExcelViewModel.Working -= SetMouseLoading;
          DefaultTaskViewModel.ChangeStep -= ChangeStep;
          DefaultTaskViewModel.Working -= SetMouseLoading;
          //Neue Events hinzufügen
          ItemSelectViewModel.ChangeStep += ChangeStep;
          ItemSelectViewModel.Working += SetMouseLoading;
          //Datenübergabe
          ItemSelectViewModel.DataItems = ExcelViewModel.DataItems;
          ItemSelectViewModel.BuildTree();
          //Alte Views abwählen und neue hinzufügen
          ExcelViewModel.ExcelViewVisible = false;
          DefaultTaskViewModel.DefaultTaskViewVisible = false;
          ItemSelectViewModel.ItemSelectViewVisible = true;
          break;
        case 3: // DefaultTaskView
          //Alte Events abwählen
          ItemSelectViewModel.ChangeStep -= ChangeStep;
          ItemSelectViewModel.Working -= SetMouseLoading;
          LoginViewModel.ChangeStep -= ChangeStep;
          LoginViewModel.Working -= SetMouseLoading;
          //Neue Events hinzufügen
          DefaultTaskViewModel.ChangeStep += ChangeStep;
          DefaultTaskViewModel.Working += SetMouseLoading;
          DefaultTaskViewModel.FilterCreatableItems(ItemSelectViewModel.DataItems, ExcelViewModel.DefaultEmployee);
          //Datenübergabe
          DefaultTaskViewModel.RefreshDefaultTaskList();
          DefaultTaskViewModel.SelectUserStories();
          //Alte Views abwählen und neue hinzufügen
          ItemSelectViewModel.ItemSelectViewVisible = false;
          LoginViewModel.LoginViewVisible = false;
          DefaultTaskViewModel.DefaultTaskViewVisible = true;
          break;
        case 4: // LoginView
          //Alte Events abwählen
          DefaultTaskViewModel.ChangeStep -= ChangeStep;
          DefaultTaskViewModel.Working -= SetMouseLoading;
          BacklogsCompareViewModel.ChangeStep -= ChangeStep;
          BacklogsCompareViewModel.Working -= SetMouseLoading;
          ProgressViewModel.ChangeStep -= ChangeStep;
          ProgressViewModel.Working -= SetMouseLoading;
          //Neue Events hinzufügen
          LoginViewModel.ChangeStep += ChangeStep;
          LoginViewModel.Working += SetMouseLoading;
          //BacklogCompare resetten
          BacklogsCompareViewModel.CompareBacklogs = false;
          //Alte Views abwählen und neue hinzufügen
          DefaultTaskViewModel.DefaultTaskViewVisible = false;
          BacklogsCompareViewModel.BacklogsCompareViewVisible = false;
          ProgressViewModel.ProgressViewVisible = false;
          LoginViewModel.LoginViewVisible = true;
          break;
        case 5: // BacklogCompareView
          //Alte Events abwählen
          LoginViewModel.ChangeStep -= ChangeStep;
          LoginViewModel.Working -= SetMouseLoading;
          DefaultTaskViewModel.ChangeStep -= ChangeStep;
          DefaultTaskViewModel.Working -= SetMouseLoading;
          //Neue Events hinzufügen
          BacklogsCompareViewModel.ChangeStep += ChangeStep;
          BacklogsCompareViewModel.Working += SetMouseLoading;
          //Datenübergabe
          BacklogsCompareViewModel.ApiConnector = LoginViewModel.ApiConnector;
          BacklogsCompareViewModel.DataItems = DefaultTaskViewModel.DataItems;
          //Alte Views abwählen und neue hinzufügen
          LoginViewModel.LoginViewVisible = false;
          ProgressViewModel.ProgressViewVisible = false;
          BacklogsCompareViewModel.BacklogsCompareViewVisible = true;
          break;
        case 6: //ProgressView
          //Alte Events abwählen
          LoginViewModel.ChangeStep -= ChangeStep;
          LoginViewModel.Working -= SetMouseLoading;
          BacklogsCompareViewModel.ChangeStep -= ChangeStep;
          BacklogsCompareViewModel.Working -= SetMouseLoading;
          //Neue Events hinzufügen
          ProgressViewModel.ChangeStep += ChangeStep;
          ProgressViewModel.Working += SetMouseLoading;
          //Datenübergabe
          ProgressViewModel.CompareBacklogs = BacklogsCompareViewModel.CompareBacklogs;
          if (ProgressViewModel.CompareBacklogs) ProgressViewModel.DataItems = BacklogsCompareViewModel.DataItems;
          else ProgressViewModel.DataItems = DefaultTaskViewModel.DataItems;
          ProgressViewModel.SortData();
          ProgressViewModel.ApiConnector = LoginViewModel.ApiConnector;
          //Alte Views abwählen und neue hinzufügen
          LoginViewModel.LoginViewVisible = false;
          BacklogsCompareViewModel.BacklogsCompareViewVisible = false;
          ProgressViewModel.ProgressViewVisible = true;
          break;

        default: break;

      }
    }
    public void SetMouseLoading(object sender, bool e)
    {
      if (e) Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait; // set the cursor to loading spinner
      else Mouse.OverrideCursor = System.Windows.Input.Cursors.Arrow; // set the cursor back to arrow
    }
  }
}
