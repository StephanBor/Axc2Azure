using AxcToAzure.Model;
using AxcToAzure.Utilities;
using AxcToAzure.View;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Resx = AxcToAzure.Properties.Resources;

namespace AxcToAzure.ViewModel
{
  public class LoginViewModel : NotifyObject
  {
    #region Properties
    public bool LoginViewVisible
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public bool UseProxy
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    
    public string Url
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public string ApiTeamId
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public string ApiProjectId
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public string Username
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public string ProxyAddress
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public APIConnector ApiConnector
    {
      get { return Get<APIConnector>(); }
      set { Set(value); }
    }

    public SecureString SecurePassword { private get; set; }

    public EventHandler<bool> Working;
    public event EventHandler<int> ChangeStep;
    #endregion
    #region Constructor
    public LoginViewModel() 
    {
      ApiProjectId = "";
      ApiTeamId = "";
      LoginViewVisible = false;
      CreateCommands();
    }
    #endregion
    #region Methods
    /// <summary>
    /// Prüft ob alle nötigen Felder (Username, Passwort,..)beschrieben sind
    /// </summary>
    /// <returns>true wenn alle Felder beschrieben sind</returns>
    public bool FieldsWritten()
    {
      if(Url == null || Url.Trim() =="")
      {
        MessageBox.Show(Resx.LoginViewModelMessageUrl, Resx.MessageWarning, MessageBoxButton.OK, MessageBoxImage.Warning);
        return false;
      }
      if (Username == null || Username.Trim() == "")
      {
        MessageBox.Show(Resx.LoginViewModelMessageUsername, Resx.MessageWarning, MessageBoxButton.OK, MessageBoxImage.Warning); ;
        return false;
      }
      if (SecurePassword == null || SecurePassword.Length == 0)
      {
        MessageBox.Show(Resx.LoginViewModelMessagePassword, Resx.MessageWarning, MessageBoxButton.OK, MessageBoxImage.Warning); ;
        return false;
      }
      if (UseProxy && (ProxyAddress == null || ProxyAddress.Trim() == ""))
      {
        MessageBox.Show(Resx.LoginViewModelMessageProxy, Resx.MessageWarning, MessageBoxButton.OK, MessageBoxImage.Warning); ;
        return false;
      }
      bool projectIdWritten = ApiProjectId != null && ApiProjectId.Trim() != "";
      bool teamIdWritten = ApiTeamId != null && ApiTeamId.Trim() != "";
      if((projectIdWritten && !teamIdWritten) || (!projectIdWritten && teamIdWritten))
      {
        MessageBox.Show(Resx.LoginViewModelMessageIds, Resx.MessageWarning, MessageBoxButton.OK, MessageBoxImage.Warning); ;
        return false;
      }
      return true;
    }
    #endregion
    #region Commands
    public ICommand ContinueCommand { get; private set; }
    public ICommand BackCommand { get; private set; }
    public ICommand CompareCommand { get; private set; }
    public ICommand OpenInstructionCommand { get; private set; }
    public void CreateCommands()
    {
      ContinueCommand = new RelayCommand(Continue);
      BackCommand = new RelayCommand(Back);
      CompareCommand = new RelayCommand(Compare);
      OpenInstructionCommand = new RelayCommand(OpenInstruction);
    }
    private void OpenInstruction()
    {
      InstructionsViewModel instructionsViewModel = new InstructionsViewModel("pack://application:,,,/Styles/Api_Instructions.png", Resx.ApiInstructionsColumn1, Resx.ApiInstructionsColumn2);
      InstructionView instructionWindow = new InstructionView();
      instructionWindow.DataContext = instructionsViewModel;
      instructionWindow.Show();
    }
    private void Continue()
    {
      if (!FieldsWritten()) return;
      Working.Invoke(this, true);
      ApiConnector = new APIConnector(Username, SecurePassword, Url, ApiTeamId, ApiProjectId, (UseProxy) ? ProxyAddress : "");
      Working.Invoke(this, false);
      ChangeStep(this, 6);
    }
    private void Compare()
    {
      if (!FieldsWritten()) return;
      Working.Invoke(this, true);
      ApiConnector = new APIConnector(Username, SecurePassword, Url, ApiTeamId, ApiProjectId, (UseProxy) ? ProxyAddress : "");
      Working.Invoke(this, false);
      ChangeStep(this, 5);
    }
    private void Back()
    {
      Working.Invoke(this, true);

      Working.Invoke(this, false);
      ChangeStep(this, 3);
    }
    
    #endregion
  }
}
