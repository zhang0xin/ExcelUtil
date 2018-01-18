using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelUtil
{
  public class ExcelReport : ExcelBase
  {
    public ExcelReport(string excelFile)
    {
      this.stream = new MemoryStream(File.ReadAllBytes(excelFile));
    }
    public ExcelReport(Stream stream)
    {
      this.stream = stream;
    }
    public byte[] ExecuteTemplate(ParameterData data)
    {
      return ExecuteTemplate(stream, data);
    }
    public static byte[] ExecuteTemplate(string excelFile, ParameterData data)
    {
      var stream = new MemoryStream(File.ReadAllBytes(excelFile));
      return ExecuteTemplate(stream, data);
    }
    public static byte[] ExecuteTemplate(Stream stream, ParameterData data)
    {
      var newStream = new MemoryStream();
      using (ExcelPackage package = new ExcelPackage(stream))
      {
        if (package != null)
        {
          foreach (var sheet in package.Workbook.Worksheets)
          {
            var addresses = GetAddressList(sheet);
            addresses.Sort(delegate(string addr1, string addr2)
            {
              return GetMaxRow(addr2).CompareTo(GetMaxRow(addr1));
            });
            foreach (string address in addresses)
            {
              ExecuteDataOnCell(sheet, address, data);
            }
          }
        }
        package.SaveAs(newStream);
      }
      newStream.Seek(0, SeekOrigin.Begin);
      return newStream.ToArray();
    }

    private static void ExecuteDataOnCell(ExcelWorksheet sheet, string address, ParameterData data)
    {
      string cellValue = DistinctValue(sheet.Cells[address].Value) + "";
      Regex regex = new Regex(@"#\{[^\}]*\}");
      if (!regex.IsMatch(cellValue)) return;

      try
      {
        bool isList = false;
        if (cellValue.StartsWith("#{"))
        {
          var config = ParameterConfig.CreateFromJson(cellValue.TrimStart('#'));
          if (config.IsSetList())
          {
            isList = true;
            InsertTableData(sheet, address, config.GetTableValue(data));
          }
        }
        if (!isList)
        {
          sheet.Cells[address].Value = regex.Replace(cellValue, delegate(Match match)
          {
            var config = ParameterConfig.CreateFromJson(match.ToString().TrimStart('#'));
            if (config.IsSetField())
            {
              return config.GetStringValue(data);
            }
            return match.ToString();
          });
        }
      }
      catch (Exception ex)
      {
        sheet.Cells[address].Value = ex.Message;
      }
    }
    static void InsertTableData(ExcelWorksheet sheet, string address, DataTable table)
    {
      var rowAddress = GetBelongsRowAddress(address);
      var indexs = rowAddress.Split(':');
      var rowStart = int.Parse(indexs[0]);
      var rowEnd = indexs.Length > 1 ? int.Parse(indexs[1]) : rowStart;
      var colStart = sheet.Cells[address].Start.Column;
      var rowCount = rowEnd - rowStart + 1;
      var copiedRowStart = rowStart + rowCount;
      var copiedRowEnd = rowEnd + rowCount;

      for (int i = 0; i < table.Rows.Count - 1; i++)
      {
        var copiedAddress = copiedRowStart + ":" + copiedRowEnd;
        sheet.InsertRow(rowEnd + 1, rowCount, rowEnd);
        sheet.Cells[rowAddress].Copy(sheet.Cells[copiedAddress]);
      }
      var enumcell = sheet.Cells[address];
      for (int i = 0, offsetRow = 0; i < table.Rows.Count; i++)
      {
        for (int j = 0, offsetCol = 0; j < table.Columns.Count; j++)
        {
          var dataCell = sheet.Cells[rowStart + offsetRow + i, colStart + offsetCol + j];
          dataCell.Value = table.Rows[i][j];
          var mergeCell = sheet.MergedCells[rowStart + offsetRow + i, colStart + offsetCol + j];
          if (dataCell.Merge)
          {
            offsetCol += sheet.Cells[mergeCell].End.Column - dataCell.End.Column;
            offsetRow += sheet.Cells[mergeCell].End.Row - dataCell.End.Row;
          }
        }
      }
    }
    static string GetBelongsRowAddress(string address)
    {
      string[] cells = address.Split(':');
      string rowAddres = "";
      if (cells.Length >= 1)
      {
        rowAddres += GetRow(cells[0]);
      }
      if (cells.Length == 2)
      {
        rowAddres += ":" + GetRow(cells[1]);
      }
      else
      {
        rowAddres += ":" + GetRow(cells[0]);
      }
      return rowAddres;
    }
    static int GetMaxRow(string address)
    {
      string[] cells = address.Split(':');
      string cell = cells.Count() > 1 ? cells[1] : cells[0];

      return GetRow(cell);
    }
    static int GetRow(string singleCellAddress)
    {
      var chars = singleCellAddress.ToCharArray();
      for (int i = 0; i < chars.Length; i++)
      {
        if (chars[i] >= '0' && chars[i] <= '9')
        {
          return int.Parse(singleCellAddress.Substring(i));
        }
      }
      return 0;
    }
  }
}
