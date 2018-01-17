using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Runtime.Serialization.Json;

namespace ExcelUtil
{
  public class ParameterData
  {
    public Dictionary<string, object> Fields {get; set;}
    public Dictionary<string, DataTable> Tables {get; set;}
    public ParameterData()
    {
      Fields = new Dictionary<string, object>();
      Tables = new Dictionary<string, DataTable>();
    }
  }
}
