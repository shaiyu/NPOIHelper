using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AspNetCoreSample.Models
{
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

        [ColumnType(Name = "用户名")]
        public string Name { get; set; }

        [ColumnType(Hide = true)]
        public string Pwd { get; set; }
    }
}
