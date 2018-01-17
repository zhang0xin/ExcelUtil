using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ExcelUtil;
using OfficeOpenXml;

namespace ExcelUtilExample
{
  class Program
  {
    static void Main(string[] args)
    {
      EPPlusHelper helper = new EPPlusHelper();
      ParameterData data1 = new ParameterData();
      data1.Fields["string_field"] = "字符串字段显示";
      data1.Fields["integer_field"] = "整数字段显示";
      data1.Tables["list1"] = CreateTestTable();
      data1.Tables["list2"] = CreateTestTable();

      ExcelPackage package1 = new ExcelPackage(new MemoryStream(helper.ExecuteTemplate("template.xlsx", data1)));
      using (Stream stream = new FileStream("out.xlsx", FileMode.Create))
      {
        package1.SaveAs(stream);
      }

      data1.Fields["weijianzhan"] = "二钢维检站";
      data1.Fields["zuoyequ"] = "作业区二";
      data1.Fields["zuoyezhang"] = "作业长1";
      data1.Fields["zhanzhang"] = "站长1";
      data1.Fields["guanliyuan"] = "管理员1";
      data1.Tables["fankuibiao"] = CreateTestTable();

      ExcelPackage package2 = new ExcelPackage(new MemoryStream(helper.ExecuteTemplate("template2.xlsx", data1)));
      using (Stream stream = new FileStream("out2.xlsx", FileMode.Create))
      {
        package2.SaveAs(stream);
      }


    }  
    public static DataTable CreateTestTable()
    {
      DataTable table = new DataTable();
      table.Columns.Add("col1");
      table.Columns.Add("col2");
      table.Columns.Add("col3");
      table.Columns.Add("col4");
      table.Columns.Add("col5");
      table.Columns.Add("col6");
      table.Columns.Add("col7");
      table.Columns.Add("col8");
      table.Columns.Add("col9");
      table.Columns.Add("col10");
      table.Columns.Add("col11");
      table.Columns.Add("col12");
      table.Columns.Add("col13");
      table.Columns.Add("col14");
      DataRow row;

      row = table.NewRow();
      row[0] = "01";
      row[1] = "val01";
      row[2] = "val01";
      row[3] = "val02";
      row[4] = "val02";
      row[5] = "val02";
      row[6] = "val02";
      row[7] = "val02";
      row[8] = "val02";
      row[9] = "val02";
      row[10] = "val02";
      row[11] = "val02";
      row[12] = "val02";
      row[13] = "val02";
      table.Rows.Add(row);

      row = table.NewRow();
      row[0] = "02";
      row[1] = "val11";
      row[2] = "val12";
      table.Rows.Add(row);

      row = table.NewRow();
      row[0] = "03";
      row[1] = "val21";
      row[2] = "val22";
      table.Rows.Add(row);

      return table;
    }

  }
}
