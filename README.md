# NPOIHelper
Export Excel NPOIHelper
# Prepare Model
        //Model
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

            [ColumnType(Name = "Test", Hide = false)]
            public string Name { get; set; }

            [ColumnType(Hide = true)]
            public string Pwd { get; set; }
        }
# 1.Instantiation
            ExcelHelper helper = new ExcelHelper();
# 2.Fill Data 
### 2.1Export List Data 
###### 2.1.1helper.Add<User>(SheetName, ListData);
            // for test List
            helper.Add<User>("使用注解的ListSheet", list);

###### 2.1.2helper.Add<User>(SheetName, ListData, Columns);
            helper.Add<User>("指定Column的ListSheet",  list,  new Column[] {
                new Column("Name", "姓名"),
                //use Fun
                new Column("Name", "姓名", ColumnType.Default, (t, index)=> {
                        var user = (User)t;
                        return user.Name + "---------" + user.Pwd;
                }),
            });

### 2.2Export DataTable Data 

###### 2.2.1helper.Add(SheetName, DataTable);
            // for test DataTable
            helper.Add("指定Column的Dt", dt, new Column[] {
                new Column("Name","姓名"),
                new Column("Pwd","密码"),
                new Column("Age","年龄", ColumnType.NumDecimal2),
                new Column("Formula", "测试公式", ColumnType.Number) { IsFormula=true },
            });

###### 2.2.2helper.Add(SheetName, DataTable, Columns);

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

# 3.Last Export
###### 3.1 Http Report
            helper.Report();
###### 3.2 Client Report
            helper.ReportClient("/test.xls");

