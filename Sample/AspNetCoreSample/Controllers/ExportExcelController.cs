using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AspNetCoreSample.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOIHelper;
using NPOIHelperWebExtension;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace AspNetCoreSample.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class ExportExcelController : ControllerBase
    {
        /// <summary>
        /// 使用Model导出报表
        /// </summary>
        /// <returns></returns>
        public IActionResult ExportExcelList()
        {
            var data = GetList();

            var helper = NPOIHelperBuild.GetHelper();
            helper.Add("用户列表", data);

            return File(helper.ToArray(), helper.ContentType, helper.FullName);
        }

        public void ExportExcelSomeList()
        {
            var data = GetList();
            var helper = NPOIHelperBuild.GetHelper(NPOIType.xls);
            helper.FileName = "List报表";
            helper.Add("用户列表1", data);
            helper.Add("用户列表2", data);
            helper.Add("用户列表3", data);

            NPOIExport.Export(Response, helper);
        }


        /// <summary>
        /// 导出DateTable
        /// </summary>
        /// <returns></returns>
        public async Task ExportExcelDataTable()
        {
            var data = GetDt();
            var helper = NPOIHelperBuild.GetHelper();
            helper.FileName = "DataTable报表";

            // for test DataTable
            helper.Add("指定Column的DataTable", data, new Column[] {
                new Column("Name","姓名"),
                new Column("Pwd","密码"),
                new Column("Age","年龄", ColumnType.NumDecimal2),
                new Column("Formula", "测试公式", ColumnType.Number) { IsFormula=true },
                new Column("BirthDay","生日", ColumnType.Date)
            });

            helper.Add("使用Fun的DataTable", data, new Column[] {
                    new Column("Name","姓名", ColumnType.Default, (t, index)=> {
                        var dr = (DataRow)t;
                        return dr["Name"]+"---------"+ dr["Pwd"] + "|Index:"+index;
                    }),
                    new Column("Age","年龄", ColumnType.NumDecimal2),
                    new Column("Age","测试公式", ColumnType.Default,(t, index)=> {
                        return "B" + index +"*B" + index;
                    }) { IsFormula=true },
            });

            Response.Clear();
            Response.ContentType = helper.ContentType;
            Response.Headers.Add("Content-Disposition", $"attachment; filename={helper.FullName};utf8");
            await Response.BodyWriter.WriteAsync(new ReadOnlyMemory<byte>(helper.ToArray()));
        }

        #region data sourrce

        public List<User> GetList()
        {
            return new List<User>() {
                new User("Labbor","123"),
                new User("Lisa"),
                new User("Lisa","sadsad"),
                new User("Lisa"),
                new User("Lisa","ppppppppp"),
            };
        }

        public static DataTable GetDt()
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
            var formula = "C" + (dt.Rows.Count + 2) + "*D" + (dt.Rows.Count + 2);
            for (int i = 0; i < 103; i++)
            {
                dt.Rows.Add("Lisa Dt" + i, "Dt" + i, i, DateTime.Now.AddDays(i), formula);
            }
            return dt;
        }
        #endregion
    }
}
