using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var list = getList();
            DataTable dt = getDt();

            //instantiation
            ExcelHelper helper = new ExcelHelper();

            // for test List
            helper.Add<User>("使用注解的List", list);

            helper.Add<User>("指定Column的List", list, new Column[] {
                new Column("Pwd","姓名"),
                new Column("Name","姓名",ColumnType.Default,(t, index)=> {
                        var user = (User)t;
                        return user.Name + "---------" + user.Pwd;
                }),
            });

            // for test DataTable
            helper.Add("指定Column的Dt", dt, new Column[] {
                new Column("Name","姓名"),
                new Column("Pwd","密码"),
                new Column("Age","年龄", ColumnType.NumDecimal2),
                new Column("Age","年龄2", ColumnType.NumDecimal2),
                new Column("Formula", "测试公式", ColumnType.Number) { IsFormula=true },
                new Column("BirthDay","生日", ColumnType.Date),
                new Column("BirthDay","生日", ColumnType.DateTime),
                new Column("BirthDay","生日", ColumnType.DateFile),
            });

            helper.Add("使用Fun的Dt", dt, new Column[] {
                    new Column("Name","姓名",ColumnType.Default,(t, index)=> {
                        var dr = (DataRow)t;
                        return dr["Name"]+"---------"+ dr["Pwd"];
                    }),
                    new Column("Age","年龄", ColumnType.NumDecimal2),
                    new Column("Age","测试公式",ColumnType.Default,(t, index)=> {
                        return "B" + index +"*B" + index;
                    }) { IsFormula=true },
            });
            // last, export
            helper.ReportClient("/test.xls");


            var file = new FileInfo("/test.xls");
            Console.WriteLine(file.FullName);
            if (System.IO.File.Exists(file.FullName))
            {
                System.Diagnostics.Process.Start(file.FullName);
            }
            Console.ReadKey();
        }

        public static List<User> getList()
        {
            return new List<User>() {
                new User("Labbor","123"),
                new User("Lisa"),
                new User("Lisa","sadsad"),
                new User("Lisa"),
                new User("Lisa","ppppppppp"),
            };
        }

        public static DataTable getDt()
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
            for (int i = 0; i < 101; i++)
            {
                dt.Rows.Add("Lisa Dt" + i, "Dt" + i, i, DateTime.Now.AddDays(i), formula);
            }
            return dt;
        }

        public class User
        {
            public User(string name)
            {
                this.Name = name;
            }

            public User(string name, string pwd)
            {
                this.Name = name;
                this.Pwd = pwd;
            }

            [ColumnType(Name = "Test")]
            public string Name { get; set; }

            [ColumnType(Hide = true)]
            public string Pwd { get; set; }
        }
    }
}
