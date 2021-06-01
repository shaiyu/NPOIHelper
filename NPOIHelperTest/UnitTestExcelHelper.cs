using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace NPOIHelperTest
{
    [TestClass]
    public class UnitTestExcelHelper : BaseUnitTest
    {
        public UnitTestExcelHelper()
        {
            if (Directory.Exists("ExportExcel"))
            {
                //删除旧数据
                Directory.Delete("ExportExcel", true);
            }
            Directory.CreateDirectory("ExportExcel");
        }

        [TestMethod]
        public void TestExportList()
        {
            var list = DataMakerHelper.Makes<ExportUser>(100).ToList();


            IExcelHelper helper = NPOIHelperBuild.GetHelper();
            helper.Add("sheet1", list);
            helper.FileName = "测试报表注解导出list";
            helper.ReportClient($"ExportExcel/{helper.FullName}");



            //判断导出
            bool exists = File.Exists($"ExportExcel/{helper.FullName}");
            Assert.IsTrue(exists);
        }

        [TestMethod]
        public void TestExportDataTable()
        {
            var dt = GetDataTable();


            IExcelHelper helper = NPOIHelperBuild.GetHelper();
            helper.Add("sheet1", dt, new Column[] {
                new Column("Name","姓名"),
                new Column("Pwd","密码"),
                new Column("Age","年龄", ColumnType.NumDecimal2),
                new Column("Age","年龄2", ColumnType.NumDecimal2),
                new Column("Formula", "测试公式", ColumnType.Number) { IsFormula=true },
                new Column("Age","测试公式2",ColumnType.Default,(t, index)=> {
                    return "C" + index +"*D" + index;
                }) { IsFormula=true },
                new Column("BirthDay","生日", ColumnType.Date),
                new Column("BirthDay","生日", ColumnType.DateTime),
                new Column("BirthDay","生日", ColumnType.DateFile),
            });

            helper.FileName = "测试报表导出DataTable";
            helper.ReportClient($"ExportExcel/{helper.FullName}");



            //判断导出
            bool exists = File.Exists($"ExportExcel/{helper.FullName}");
            Assert.IsTrue(exists);
        }


        public static DataTable GetDataTable(int count = 100)
        {
            var dt = new DataTable();
            DataColumn colName = new DataColumn("Name", typeof(string));
            DataColumn colPwd = new DataColumn("Pwd", typeof(string));
            DataColumn colAge = new DataColumn("Age", typeof(int));
            DataColumn colFormula = new DataColumn("Formula", typeof(string));
            DataColumn colBirthDay = new DataColumn("BirthDay", typeof(DateTime));
            dt.Columns.Add(colName);
            dt.Columns.Add(colPwd);
            dt.Columns.Add(colAge);
            dt.Columns.Add(colBirthDay);
            dt.Columns.Add(colFormula);

            //dt.Rows.Count + 1 是当前行  再+1, 是即将要加的行
            for (int i = 0; i < count; i++)
            {
                var formula = "C" + (i + 2) + "*D" + (i + 2);
                dt.Rows.Add("Lisa Dt" + i, "Dt" + i, i, DateTime.Now.AddDays(i), formula);
            }
            return dt;
        }
    }
}
