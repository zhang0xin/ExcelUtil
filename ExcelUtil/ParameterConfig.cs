using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.IO;

namespace ExcelUtil
{
  public class ParameterConfig
  {
    public string DataField { get; set; }
    public string DataList { get; set; }

    public static ParameterConfig CreateFromJson(string json)
    {
      JavaScriptSerializer js = new JavaScriptSerializer();
      return js.Deserialize<ParameterConfig>(json);
    }
    public bool IsSetField()
    {
      return !string.IsNullOrWhiteSpace(DataField);
    }
    public bool IsSetList()
    { 
      return !string.IsNullOrWhiteSpace(DataList);
    }
    public object GetFieldValue(ParameterData data)
    {
      if (!data.Fields.Keys.Contains(DataField))
        return string.Format("Error: {0} field config not exist", DataField);
      return data.Fields[DataField];
    }
    public int GetIntValue(ParameterData data)
    {
      return int.Parse(GetFieldValue(data)+"");
    }
    public string GetStringValue(ParameterData data)
    {
      return GetFieldValue(data) + "";
    }
    public DataTable GetTableValue(ParameterData data)
    {
      if (!data.Tables.Keys.Contains(DataList))
        return null;
      return data.Tables[DataList];
    }
  }
}
