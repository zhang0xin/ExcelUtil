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
  public class EPPlusHelper
  {
    public byte[] ExecuteTemplate(string excelFile , ParameterData data)
    {
      var stream = new MemoryStream(File.ReadAllBytes(excelFile));
      return ExecuteTemplate(stream, data);
    }
    public byte[] ExecuteTemplate(Stream stream, ParameterData data)
    {
      var newStream = new MemoryStream();
      using(ExcelPackage package = new ExcelPackage(stream))
      {
        if (package != null)
        {
          foreach (var sheet in package.Workbook.Worksheets)
          {
            var addresses = GetAddressList(sheet);
            addresses.Sort(delegate(string addr1, string addr2) {
              return GetMaxRow(addr2).CompareTo(GetMaxRow(addr1));
            });
            foreach(string address in addresses)
            {
              string cellValue = DistinctValue(sheet.Cells[address].Value)+"";
              Regex regex = new Regex(@"#\{[^\}]*\}");
              if (regex.IsMatch(cellValue))
              {
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
                    sheet.Cells[address].Value = regex.Replace(cellValue, delegate(Match match) {
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
            }
          }
        }
        package.SaveAs(newStream);
      }
      newStream.Seek(0, SeekOrigin.Begin);
      return newStream.ToArray();
    }
    public void InsertTableData(ExcelWorksheet sheet, string address, DataTable table)
    {
      var rowAddress = GetBelongsRowAddress(address);
      var indexs = rowAddress.Split(':');
      var rowStart = int.Parse(indexs[0]);
      var rowEnd = indexs.Length>1? int.Parse(indexs[1]) : rowStart;
      var colStart = sheet.Cells[address].Start.Column;
      var rowCount = rowEnd - rowStart + 1;
      var copiedRowStart = rowStart + rowCount;
      var copiedRowEnd = rowEnd + rowCount;

      for (int i = 0; i < table.Rows.Count-1; i++)
      {
        var copiedAddress = copiedRowStart + ":" + copiedRowEnd;
        sheet.InsertRow(rowEnd + 1, rowCount, rowEnd);
        sheet.Cells[rowAddress].Copy(sheet.Cells[copiedAddress]);
      }
      var enumcell = sheet.Cells[address];
      for (int i=0, offsetRow=0; i<table.Rows.Count; i++)
      {
        for (int j = 0, offsetCol = 0; j < table.Columns.Count; j++)
        {
          var dataCell = sheet.Cells[rowStart+offsetRow+i, colStart+offsetCol+j];
          dataCell.Value = table.Rows[i][j];
          var mergeCell = sheet.MergedCells[rowStart+offsetRow+i, colStart+offsetCol+j];
          if(dataCell.Merge)
          {
            //offset += GetColNextMergeCount(sheet, dataCell);
            offsetCol += sheet.Cells[mergeCell].End.Column - dataCell.End.Column;
            offsetRow += sheet.Cells[mergeCell].End.Row - dataCell.End.Row;
          }
        }
      }
    }
    public int GetColNextMergeCount(ExcelWorksheet sheet, ExcelAddress cell)
    {
      var nextCell = sheet.Cells[cell.Start.Row, cell.Start.Column + 1];
      var count = 0;
      while (nextCell.Merge)
      {
        count ++;
        nextCell = sheet.Cells[nextCell.Start.Row, nextCell.Start.Column + 1];
      }
      return count;
    }
    public string GenerateExecutableCode(string excelFile)
    {
      var stream = new MemoryStream(File.ReadAllBytes(excelFile));
      return GenerateExecutableCode(stream);
    }
    public string GenerateExecutableCode(Stream stream)
    {
      return
      @"using System;
        using OfficeOpenXml;
        using OfficeOpenXml.Style;
        using System.IO;
        using System.Drawing;

        namespace DynamicCodes{
          class CodeWrapper {
            public static void Execute(){
              using (ExcelPackage package = new ExcelPackage()){
                "+GenerateCode(stream)+@"
                using (Stream stream = new FileStream(""./output.xlsx"", FileMode.Create)){
                  package.SaveAs(stream);
                }
              }
            }
          }
        }
      ";
    }
    public string GenerateCode(string excelFile)
    {
      var stream = new MemoryStream(File.ReadAllBytes(excelFile));
      return GenerateCode(stream);
    }
    public string GenerateCode(Stream stream)
    {
      StringBuilder code = new StringBuilder();
      using(ExcelPackage package = new ExcelPackage(stream))
      {
        if (package != null)
        {
          code.AppendLine("ExcelWorksheet sheet;");
          foreach(var sheet in package.Workbook.Worksheets)
          {
            string codeCreateSheet =
              "sheet = package.Workbook.Worksheets.Add(\"{0}\");";
            code.AppendLine(string.Format(codeCreateSheet, sheet.Name));
            for(int i=sheet.Dimension.Start.Row; i<=sheet.Dimension.End.Row; i++)
            {
              string codeSetHeight = "sheet.Row({0}).Height = {1};";
              code.AppendLine(string.Format(codeSetHeight, i, sheet.Row(i).Height));
            }
            for(int i=sheet.Dimension.Start.Column; i<=sheet.Dimension.End.Column; i++)
            {
              string codeSetWidth = "sheet.Column({0}).Width = {1};";
              code.AppendLine(string.Format(codeSetWidth, i, sheet.Column(i).Width));
            }
            foreach(string address in GetAddressList(sheet))
            {
              string codeSetValue = "sheet.Cells[\"{0}\"].Value = \"{1}\";";
              string cellValue = DistinctValue(sheet.Cells[address].Value)+"";
              code.AppendLine(GenerateCellStyleCodes(sheet.Cells[address]));
              if (!string.IsNullOrEmpty(cellValue))
              {
                if (sheet.MergedCells.Contains(address))
                {
                  string codeSetMerged = "sheet.Cells[\"{0}\"].Merge = true;";
                  code.AppendLine(string.Format(codeSetMerged, address));
                }
                code.AppendLine(string.Format(codeSetValue, address, EncodeCodeString(cellValue)));
              }
            }
          }
        }
      }
      return code.ToString();
    }
    public string GenerateRowStyleCodes(ExcelRow row)
    {
      return GenerateStyleCodes("sheet.Row("+row.Row+")", row.Style);
    }
    public string GenerateCellStyleCodes(ExcelRange range)
    {
      return GenerateStyleCodes("sheet.Cells[\"" + range.Address + "\"]", range.Style);
    }
    public string GenerateStyleCodes(string stylePrefix, ExcelStyle style)
    {
      StringBuilder codes = new StringBuilder();
      string codeFormat;
      codeFormat = "{0}.Style.Border.Left.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Border.Left.Style));

      codeFormat = "{0}.Style.Border.Right.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Border.Right.Style));

      codeFormat = "{0}.Style.Border.Top.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Border.Top.Style));

      codeFormat = "{0}.Style.Border.Bottom.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Border.Bottom.Style));

      codeFormat = "{0}.Style.Numberformat.Format = \"{1}\";";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, EncodeCodeString(style.Numberformat.Format)));

      codeFormat = "{0}.Style.Font.Bold = {1};";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Font.Bold.ToString().ToLower()));

      codeFormat = "{0}.Style.Font.Size = {1};";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Font.Size));

      codeFormat = "{0}.Style.Font.Name = \"{1}\";";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Font.Name));

      codeFormat = "{0}.Style.Font.Color.SetColor(Color.FromArgb({1}));";
      if (!string.IsNullOrEmpty(style.Font.Color.Rgb))
      {
        codes.AppendLine(string.Format(
          codeFormat, stylePrefix, RgbToParameters(style.Font.Color.Rgb)));
      }

      codeFormat = "{0}.Style.Fill.PatternType = "+
          " (ExcelFillStyle) Enum.Parse(typeof(ExcelFillStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Fill.PatternType));
      if (!string.IsNullOrEmpty(style.Fill.BackgroundColor.Rgb))
      {
        codeFormat = "{0}.Style.Fill.BackgroundColor.SetColor(Color.FromArgb({1}));";
        codes.AppendLine(string.Format(
          codeFormat, stylePrefix, RgbToParameters(style.Fill.BackgroundColor.Rgb)));
      }

      codeFormat = "{0}.Style.VerticalAlignment = "+
          " (ExcelVerticalAlignment) Enum.Parse(typeof(ExcelVerticalAlignment), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.VerticalAlignment));

      codeFormat = "{0}.Style.HorizontalAlignment = "+
          " (ExcelHorizontalAlignment) Enum.Parse(typeof(ExcelHorizontalAlignment), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.HorizontalAlignment));

      codeFormat = "{0}.Style.WrapText = {1};";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.WrapText.ToString().ToLower()));

      codeFormat = "{0}.Style.ReadingOrder = "+
          " (ExcelReadingOrder) Enum.Parse(typeof(ExcelReadingOrder), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.ReadingOrder));

      codeFormat = "{0}.Style.WrapText = {1};";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.WrapText.ToString().ToLower()));

      codeFormat = "{0}.Style.ShrinkToFit = {1};";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.ShrinkToFit.ToString().ToLower()));

      codeFormat = "{0}.Style.Indent = {1};";
      codes.AppendLine(string.Format(
        codeFormat, stylePrefix, style.Indent));

      return codes.ToString();
    }
    public string RgbToParameters(string rgb)
    {
      //AARRGGBB -> 0xAA, 0xRR, 0xGG, 0xBB
      return rgb.Insert(6, ", 0x").Insert(4, ", 0x").Insert(2, ", 0x").Insert(0, "0x");
    }
    public string EncodeCodeString(string codes)
    {
      return codes.Replace(@"\", @"\\").Replace("\"", "\\\"");
    }
    public List<string> GetAddressList(ExcelWorksheet sheet)
    {
      List<string> addressList = new List<string>();
      foreach(var address in sheet.MergedCells)
      {
        addressList.Add(address);
      }
      foreach(var range in sheet.Cells)
      {
        if (range.Merge) continue;
        addressList.Add(range.Address);
      }
      return addressList;
    }
    public string GetBelongsRowAddress(string address)
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
    public int GetMaxRow(string address)
    {
      string [] cells = address.Split(':');
      string cell = cells.Count() > 1? cells[1] : cells[0];

      return GetRow(cell);
    }
    public int GetRow(string singleCellAddress)
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
    public object DistinctValue(object val)
    {
      var arr = val as object[,];
      if (arr != null && arr.GetLength(0) > 0 && arr.GetLength(1) > 0)
        return arr[0,0];
      else
        return val;
    }
  }
}
