using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelperTest
{
    public class ExportUser
    {
        [ColumnType(Name = "ID")]
        public Guid Id { get; set; }

        [ColumnType(Hide = true)]
        public string Pwd { get; set; }

        [ColumnType(Name = "姓名")]
        public string Name { get; set; }

        [ColumnType(Name = "年龄", Type = ColumnType.Number, IsZeroFillEmpty = true)]
        public int Age { get; set; }

        [ColumnType(Name = "手机号码")]
        public string Phone { get; set; }

        [ColumnType(Name = "出生日期", Type = ColumnType.Date)]
        public DateTime Birthday { get; set; }

        [ColumnType(Name = "身高(米)", Type = ColumnType.NumDecimal2)]
        public decimal Height { get; set; }

        [ColumnType(Name = "登记日期", Type = ColumnType.DateTime)]
        public DateTime CreateTime { get; set; }
    }
}
