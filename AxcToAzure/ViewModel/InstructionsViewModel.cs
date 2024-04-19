using AxcToAzure.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AxcToAzure.ViewModel
{
  public class InstructionsViewModel : NotifyObject
  {
    #region Properties
    public string Source
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public string Column1
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    public string Column2
    {
      get { return Get<string>(); }
      set { Set(value); }
    }
    #endregion
    #region Constructor
    public InstructionsViewModel(string source, string coulumn1, string column2) 
    { 
      Source =source; 
      Column1 = coulumn1; 
      Column2 = column2;
    }
    #endregion
    #region Methods

    #endregion
    #region Commands

    #endregion
  }

}
