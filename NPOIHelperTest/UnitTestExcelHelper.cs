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
                //ɾ��������
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
            helper.FileName = "���Ա���ע�⵼��list";
            helper.ReportClient($"ExportExcel/{helper.FullName}");



            //�жϵ���
            bool exists = File.Exists($"ExportExcel/{helper.FullName}");
            Assert.IsTrue(exists);
        }

        [TestMethod]
        public void TestExportDataTable()
        {
            var dt = GetDataTable();


            IExcelHelper helper = NPOIHelperBuild.GetHelper();
            helper.Add("sheet1", dt, new Column[] {
                new Column("Name","����"),
                new Column("Pwd","����"),
                new Column("Age","����", ColumnType.NumDecimal2),
                new Column("Age","����2", ColumnType.NumDecimal2),
                new Column("Formula", "���Թ�ʽ", ColumnType.Number) { IsFormula=true },
                new Column("Age","���Թ�ʽ2",ColumnType.Default,(t, index)=> {
                    return "C" + index +"*D" + index;
                }) { IsFormula=true },
                new Column("BirthDay","����", ColumnType.Date),
                new Column("BirthDay","����", ColumnType.DateTime),
                new Column("BirthDay","����", ColumnType.DateFile),
            });

            helper.FileName = "���Ա�����DataTable";
            helper.ReportClient($"ExportExcel/{helper.FullName}");



            //�жϵ���
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

            //dt.Rows.Count + 1 �ǵ�ǰ��  ��+1, �Ǽ���Ҫ�ӵ���
            for (int i = 0; i < count; i++)
            {
                var formula = "C" + (i + 2) + "*D" + (i + 2);
                dt.Rows.Add("Lisa Dt" + i, "Dt" + i, i, DateTime.Now.AddDays(i), formula);
            }
            return dt;
        }
    }
}
