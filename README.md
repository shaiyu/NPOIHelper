支持netstandard2.0, net45以上
# NPOIHelper
Export Excel NPOIHelper   
use 
https://github.com/shaiyu/NPOIHelper
# 1.使用List-Model导出报表
######使用注解 ColumnType
Name：导出的标题名称
Hide：是否不导出，默认导出 
Type: 导出的类型
```
        //Model
        public class User
        {
            [ColumnType(Name = "Test", Hide = false, Type = ColumnType.String)]
            public string Name { get; set; }

            [ColumnType(Hide = true)]
            public string Pwd { get; set; }
        }
```
```
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
```
# 2.使用List-Dynamic导出报表
```
        helper.Add<User>("指定Column导出", list, new Column[] {
            new Column("Name", "姓名"),
            //use Fun
            new Column("Name", "姓名", ColumnType.Default, (t, index)=> {
                    var user = (User)t;
                    return user.Name + "---------" + user.Pwd;
            }),
        });
```
# 3 使用DateTable导出报表
```
        helper.Add("使用Fun的Dt", dt, new Column[] {
                new Column("Name", "姓名", ColumnType.Default, (t, index)=> {
                    var dr = (DataRow)t;
                    return dr["Name"]+"---------"+ dr["Pwd"];
                }),
                new Column("Age","年龄", ColumnType.NumDecimal2),
                new Column("Age","测试公式",ColumnType.Default,(t, index)=> {
                    return "B" + index +"*B" + index;
                }) { IsFormula=true },
        });
```