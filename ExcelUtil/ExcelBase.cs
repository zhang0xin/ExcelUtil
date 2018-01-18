using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace ExcelUtil
{
  public class ExcelBase
  {
    protected Stream stream;
    protected static List<string> GetAddressList(ExcelWorksheet sheet)
    {
      List<string> addressList = new List<string>();
      foreach(var range in sheet.Cells)
      {
        if (range.Merge)
        {
          var address = sheet.MergedCells[range.Start.Row, range.Start.Column];
          if(!addressList.Contains(address))
            addressList.Add(sheet.MergedCells[range.Start.Row, range.Start.Column]);
        }
        else
        {
          addressList.Add(range.Address);
        }
      }
      return addressList;
    }
    protected static object DistinctValue(object val)
    {
      var arr = val as object[,];
      if (arr != null && arr.GetLength(0) > 0 && arr.GetLength(1) > 0)
        return arr[0,0];
      else
        return val;
    }
  }
}
