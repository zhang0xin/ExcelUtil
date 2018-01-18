using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelUtil;
using OfficeOpenXml;
using System.IO;
using System.Data;

namespace ExcelUtilTest
{
  [TestClass]
  public class MainTest
  {
    [TestMethod]
    public void TestExecuteTemplate()
    {
      int byteLength = 5000;
      var buffer = new byte[byteLength];
      using (ExcelPackage package = new ExcelPackage())
      {
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("for test");
        sheet.Cells["B1"].Value = @"#{DataList: 'list1'}";
        sheet.Cells["B2"].Value = @"#{DataField: 'string_field'}";
        sheet.Cells["B3"].Value = @"#{DataField: 'integer_field'}";

        sheet.Cells["C4"].Value = @"#{DataList: 'list2'}";
        sheet.Cells["C5"].Value = @"#{DataField: 'no_field'}";
        sheet.Cells["C6"].Value = @"#{field: gramma error'}";
        package.SaveAs(new MemoryStream(buffer));
      }

      var helper = new ExcelReport(new MemoryStream(buffer));
      var data = new ParameterData();
      data.Fields["string_field"] = "string field";
      data.Fields["integer_field"] = 99;
      data.Tables["list1"] = CreateTestTable();
      data.Tables["list2"] = CreateTestTable();
      var newBuffer = helper.ExecuteTemplate(data);

      using (ExcelPackage package = new ExcelPackage(new MemoryStream(newBuffer)))
      {
        var cells = package.Workbook.Worksheets["for test"].Cells;
        Assert.AreEqual("string field", cells["B2"].Value);
        Assert.AreEqual("99", cells["B3"].Value);
        Assert.AreEqual("Error: no_field field config not exist", cells["C5"].Value);
        Assert.AreEqual("无效的 JSON 基元: gramma。", cells["C6"].Value);
      }
    }
    public DataTable CreateTestTable()
    {
      DataTable table = new DataTable();
      table.Columns.Add("col1");
      table.Columns.Add("col2");
      table.Columns.Add("col3");
      DataRow row;
      row = table.NewRow();
      row[0] = "val00";
      row[1] = "val01";
      row[2] = "val02";
      row = table.NewRow();
      row[0] = "val10";
      row[1] = "val11";
      row[2] = "val12";
      row = table.NewRow();
      row[0] = "val20";
      row[1] = "val21";
      row[2] = "val22";
      return table;
    }
    [TestMethod]
    public void TestLoadFromJson()
    {
      string json = @"{""DataField"": ""data_field_name"", ""DataList"": ""data_list_name""}";
      ParameterConfig config = ParameterConfig.CreateFromJson(json);
      Assert.AreEqual("data_field_name", config.DataField);
      
      json = @"{DataField: ""data_field_name"", DataList: ""data_list_name""}";
      config = ParameterConfig.CreateFromJson(json);
      Assert.AreEqual("data_list_name", config.DataList);
    }
  }
}
