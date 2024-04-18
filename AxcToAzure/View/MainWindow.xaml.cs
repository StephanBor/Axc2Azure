using AxcToAzure.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AxcToAzure
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    public MainWindow()
    {
      var Result = MessageBox.Show(AxcToAzure.Properties.Resources.MessageSetLanguage, AxcToAzure.Properties.Resources.MessageSetLanguageHeader, MessageBoxButton.YesNo, MessageBoxImage.Question);

      if (Result == MessageBoxResult.Yes) System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("de-DE");
      else System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en");
      DataContext = new MainWindowViewModel();
      InitializeComponent();
    }
  }
}
