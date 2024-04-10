using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;

namespace AxcToAzure.Utilities
{
  public interface IRelayCommand : ICommand
  {
    void UpdateCanExecuteState();
  }

  /// <summary>
  /// Ein Kommando Objekt fuer Delegaten
  /// </summary>
  /// <typeparam name="T"></typeparam>
  public class RelayCommand<T> : IRelayCommand
  {
    #region Konstrutkor
    //ToDo: Discuss FleetManagement
    readonly Func<T, bool> _canExecuteParam;
    readonly Func<bool> _canExecute;
    readonly Action<T> _execute;

    /// <summary>
    /// Initialisert eine neue Instance von <see cref="RelayCommand&lt;T&gt;"/> Klasse, die immer ausgefuehrt werden kann
    /// </summary>
    /// <param name="execute">Auszuführende Logik</param>
    public RelayCommand(Action<T> execute)
    {
      if (execute == null)
        throw new ArgumentNullException("execute");
      _execute = execute;
    }

    /// <summary>
    /// Initialisert eine neue Instance von <see cref="RelayCommand&lt;T&gt;"/> Klasse
    /// </summary>
    /// <param name="execute">Auszuführende Logik</param>
    /// <param name="canExecute">Status der auszuführenden Logik</param>
    /// 
    public RelayCommand(Action<T> execute, Func<T, bool> canExecute)
    {
      if (execute == null)
        throw new ArgumentNullException("execute");
      _execute = execute;
      _canExecuteParam = canExecute;
    }



    public RelayCommand(Action<T> execute, Func<bool> canExecute)
    {
      if (execute == null)
        throw new ArgumentNullException("execute");
      _execute = execute;
      _canExecute = canExecute;
    }

    #endregion

    #region IRelayCommand Mebers
    public event EventHandler CanExecuteChanged;

    public void UpdateCanExecuteState()
    {
      if (CanExecuteChanged != null)
        CanExecuteChanged(this, EventArgs.Empty);
    }

    [DebuggerStepThrough]
    public bool CanExecute(object parameter)
    {
      if (parameter != null)
        return _canExecuteParam == null ? true : _canExecuteParam((T)parameter);

      return _canExecute == null ? true : _canExecute();
    }

    public void Execute(object parameter)
    {
      _execute((T)parameter);
    }
    #endregion
  }

  public class RelayCommand : IRelayCommand
  {
    #region Konstrutkor
    readonly Func<bool> _canExecute;
    readonly Action _execute;

    /// <summary>
    /// Initialisert eine neue Instance von <see cref="RelayCommand"/> Klasse, die immer ausgefuehrt werden kann
    /// </summary>
    /// <param name="execute">Auszuführende Logik</param>
    public RelayCommand(Action execute)
      : this(execute, null) { }

    /// <summary>
    /// Initialisert eine neue Instance von <see cref="RelayCommand"/> Klasse
    /// </summary>
    /// <param name="execute">Auszuführende Logik</param>
    /// <param name="canExecute">Status der auszuführenden Logik</param>
    public RelayCommand(Action execute, Func<bool> canExecute)
    {
      if (execute == null)
        throw new ArgumentNullException("execute");
      _execute = execute;
      _canExecute = canExecute;
    }
    #endregion

    #region IRelayCommand Mebers
    public event EventHandler CanExecuteChanged;

    public void UpdateCanExecuteState()
    {
      if (CanExecuteChanged != null)
        CanExecuteChanged(this, EventArgs.Empty);
    }

    [DebuggerStepThrough]
    public bool CanExecute(object parameter)
    {
      return _canExecute == null ? true : _canExecute();
    }

    public void Execute(object parameter)
    {
      _execute();
    }
    #endregion
  }

  /// <summary>
  /// Klasse um die CommandParameter zu binden und nach dem Binden ein explizites
  /// UpdateCanExecuteState auf dem Command aufrufen
  /// </summary>
  public static class BoundRelayCommand
  {
    /// <summary>
    /// Erweiterungsmoeglichkeiten um noch weitere Klasen zu unterstuetzen
    /// </summary>
    public static PropertyChangedCallback ExtendedParameterChanged { get; set; }

    /// <summary>
    /// Parameter an das Angehaengte Control weitergeben, und dann ein UpdateNotify auf dem Command-Objekt rufen
    /// </summary>
    /// <param name="d"></param>
    /// <param name="e"></param>
    public static void ParameterChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
      if (d is System.Windows.Controls.Primitives.ButtonBase)
      {
        var btn = d as System.Windows.Controls.Primitives.ButtonBase;
        btn.CommandParameter = e.NewValue;
        var relayCmd = btn.Command as IRelayCommand;
        if (relayCmd != null)
          relayCmd.UpdateCanExecuteState();
        return;
      }

      if (d is System.Windows.Controls.MenuItem)
      {
        var mi = d as System.Windows.Controls.MenuItem;
        mi.CommandParameter = e.NewValue;
        var relayCmd = mi.Command as IRelayCommand;
        if (relayCmd != null)
          relayCmd.UpdateCanExecuteState();
        return;
      }

      if (d is System.Windows.Documents.Hyperlink)
      {
        var hl = d as System.Windows.Documents.Hyperlink;
        hl.CommandParameter = e.NewValue;
        var relayCmd = hl.Command as IRelayCommand;
        if (relayCmd != null)
          relayCmd.UpdateCanExecuteState();
        return;
      }

      if (d is InputBinding)
      {
        var ib = d as InputBinding;
        ib.CommandParameter = e.NewValue;
        var relayCmd = ib.Command as IRelayCommand;
        if (relayCmd != null)
          relayCmd.UpdateCanExecuteState();
        return;
      }

      if (d is System.Windows.Shell.ThumbButtonInfo)
      {
        var tbi = d as System.Windows.Shell.ThumbButtonInfo;
        tbi.CommandParameter = e.NewValue;
        var relayCmd = tbi.Command as IRelayCommand;
        if (relayCmd != null)
          relayCmd.UpdateCanExecuteState();
        return;
      }

      // Wenn eine Extendend-Prozedur eingehaengt wurde, dann diese auch noch abarbeiten
      if (ExtendedParameterChanged != null)
        ExtendedParameterChanged(d, e);
    }

    public static readonly DependencyProperty ParameterProperty = DependencyProperty.RegisterAttached(
      "Parameter", typeof(object), typeof(BoundRelayCommand), new UIPropertyMetadata(null, ParameterChanged));

    public static object GetParameter(DependencyObject d)
    {
      return d.GetValue(ParameterProperty);
    }

    public static void SetParameter(DependencyObject d, object value)
    {
      d.SetValue(ParameterProperty, value);
    }
  }
}
