/*
 * Created by SharpDevelop.
 * User: 123
 * Date: 2018-01-08
 * Time: 10:12
 *
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Reflection;
using Microsoft.CSharp;
using System.CodeDom;
using System.CodeDom.Compiler;

namespace ExcelUtil
{
  class Program
  {
    public static void Main(string[] args)
    {
      EPPlusHelper eppHelper = new EPPlusHelper();
      string codes = eppHelper.GenerateExecutableCode("template.xlsx");

      StreamWriter writer = new StreamWriter("CodeWrapper.cs");
      writer.Write(codes);
      writer.Close();

      ExecuteCodes(codes);
      //MakeExampleExcel();
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
        //Console.WriteLine(wrapperCodes);
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
    public static void MakeExampleExcel(){
      using (ExcelPackage package = new ExcelPackage())
      {
          ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
          sheet.Cells["a1:m1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["a1:m1"].Merge = true;
          sheet.Cells["a1:m1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
          sheet.Cells["a1:m1"].Value = "自动化系统检修计划表2";
          sheet.Cells["a2:m2"].Merge = true;
          sheet.Cells["a2:m2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["a2:m2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
          sheet.Cells["a2:m2"].Value = "编号：GLWJ-ZDH-06-1712-01";
          sheet.Cells["a3:c3"].Merge = true;
          sheet.Cells["a3:c3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["a3:c3"].Value = "维检站：";
          sheet.Cells["e3:j3"].Merge = true;
          sheet.Cells["e3:j3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["e3:j3"].Value = "作业区：";
          sheet.Cells["d3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["k3:m3"].Merge = true;
          sheet.Cells["k3:m3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["k3:m3"].Value = "检修日期：                检修时间：       至";

          sheet.Cells["a4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["a4"].Value = "序号";
          sheet.Cells["b4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["b4"].Value = "检修区域或设备名称";
          sheet.Cells["c4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["c4"].Value = "检修项目和内容";
          sheet.Cells["d4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["d4"].Value = "检修项目专业";
          sheet.Cells["e4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["e4"].Value = "检修项目性质";
          sheet.Cells["f4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["f4"].Value = "项目来源描述";
          sheet.Cells["g4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["g4"].Value = "计划工时(人时)";
          sheet.Cells["h4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["h4"].Value = "技术要求";
          sheet.Cells["i4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["i4"].Value = "点检员及联系方式";
          sheet.Cells["j4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["j4"].Value = "危险预知";
          sheet.Cells["k4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["k4"].Value = "安全措施";
          sheet.Cells["l4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["l4"].Value = "施工负责人";
          sheet.Cells["m4"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
          sheet.Cells["m4"].Value = "备注";

          using (Stream stream = new FileStream("./example.xlsx", FileMode.Create))
          {
              package.SaveAs(stream);
          }
      }
    }
  }
}
