using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.CSharp;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelUtil
{
  public class ExcelCodes : ExcelBase
  {
    public ExcelCodes(string excelFile)
    {
      this.stream = new MemoryStream(File.ReadAllBytes(excelFile));
    }
    public ExcelCodes(Stream stream)
    {
      this.stream = stream;
    }
    public string GenerateExecutableCode()
    {
      return GenerateExecutableCode(this.stream);
    }
    public string GenerateCode()
    {
      return GenerateCode(stream);
    }

    public static void ExecuteCodes(string codes){
      CompilerParameters parameters = new CompilerParameters();
      parameters.ReferencedAssemblies.Add("System.dll");
      parameters.ReferencedAssemblies.Add("EPPlus.dll");
      parameters.ReferencedAssemblies.Add("System.Drawing.dll");
      parameters.GenerateExecutable = false;
      parameters.GenerateInMemory = true;
      CompilerResults results =
        (new CSharpCodeProvider()).CompileAssemblyFromSource(parameters, codes);
      if (results.Errors.HasErrors)
      {
        foreach(CompilerError error in results.Errors)
        {
          Console.WriteLine(error.ToString());
        }
      }
      else
      {
        Type type = results.CompiledAssembly.GetType("DynamicCodes.CodeWrapper");
        type.GetMethod("Execute").Invoke(null, null);
      }
    }
    public static string GenerateExecutableCode(string excelFile)
    {
      var stream = new MemoryStream(File.ReadAllBytes(excelFile));
      return GenerateExecutableCode(stream);
    }
    public static string GenerateExecutableCode(Stream stream)
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
    public static string GenerateCode(string excelFile)
    {
      var stream = new MemoryStream(File.ReadAllBytes(excelFile));
      return GenerateCode(stream);
    }
    public static string GenerateCode(Stream stream)
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

            if (sheet.Dimension == null) continue;
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
    static string GenerateCellStyleCodes(ExcelRange range)
    {
      return GenerateStyleCodes("sheet.Cells[\"" + range.Address + "\"]", range.Style);
    }
    static string GenerateStyleCodes(string stylePrefix, ExcelStyle style)
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
    static string RgbToParameters(string rgb)
    {
      //AARRGGBB -> 0xAA, 0xRR, 0xGG, 0xBB
      return rgb.Insert(6, ", 0x").Insert(4, ", 0x").Insert(2, ", 0x").Insert(0, "0x");
    }
    static string EncodeCodeString(string codes)
    {
      return codes.Replace(@"\", @"\\").Replace("\"", "\\\"");
    }
  }
}
