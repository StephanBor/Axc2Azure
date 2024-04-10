using AxcToAzure.Model;
using AxcToAzure.Utilities;
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
using xls2aturenet6.Model;

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
    public string ApiParam1
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public string ApiParam2
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
      LoginViewVisible = false;
      CreateCommands();
    }
    #endregion
    #region Methods
    public bool FieldsWritten()
    {
      if(Url == null || Url.Trim() =="")
      {
        MessageBox.Show("Please enter a valid URL!");
        return false;
      }
      if (Username == null || Username.Trim() == "")
      {
        MessageBox.Show("Please enter a valid Usename!");
        return false;
      }
      if (SecurePassword == null || SecurePassword.Length == 0)
      {
        MessageBox.Show("Please enter a valid Password!");
        return false;
      }
      if (UseProxy && (ProxyAddress == null || ProxyAddress.Trim() == ""))
      {
        MessageBox.Show("Please enter a valid Proxy Adress or uncheck the Box!");
        return false;
      }
      return true;
    }
    #endregion
    #region Commands
    public ICommand ContinueCommand { get; private set; }
    public ICommand BackCommand { get; private set; }
    public ICommand CompareCommand { get; private set; }
    public void CreateCommands()
    {
      ContinueCommand = new RelayCommand(Continue);
      BackCommand = new RelayCommand(Back);
      CompareCommand = new RelayCommand(Compare);
    }
    private void Continue()
    {
      if (!FieldsWritten()) return;
      Working.Invoke(this, true);
      ApiConnector = new APIConnector(Username, SecurePassword, Url, (UseProxy) ? ProxyAddress : "");
      Working.Invoke(this, false);
      ChangeStep(this, 6);
    }
    private void Compare()
    {
      if (!FieldsWritten()) return;
      Working.Invoke(this, true);
      ApiConnector = new APIConnector(Username, SecurePassword, Url, (UseProxy) ? ProxyAddress : "");
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
