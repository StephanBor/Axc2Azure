using AxcToAzure.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AxcToAzure.Model
{
  public class DataItem : NotifyObject
  {
    //For Program
    public string Id { get; set; }
    public string Name { get; set; }
    public string Type { get; set; }
    public string ParentId { get; set; }
    public bool CreateThis
    {
      get { return Get<bool>(); }
      set { Set(value); }
    }
    public string Employee { get; set; }
    public string DefaultTask { get; set; }
    public ObservableCollection<DataItem> Children { get; set; }

    //For Azure
    public int AzureId { get; set; }
    public string AzureDate { get; set; }
    public string AzureEmployee{get; set;}
    public DataItem() {
    Children = new ObservableCollection<DataItem>();
      AzureEmployee = "";
      CreateThis = true;
    }
    
  }
}

