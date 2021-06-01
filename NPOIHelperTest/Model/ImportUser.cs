using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelperTest
{
    public class ImportUser
    {
        [TitleName("ID")]
        public Guid Id { get; set; }

        [TitleName("姓名")]
        public string Name { get; set; }

        [TitleName("年龄")]
        public int Age { get; set; }

        [TitleName("手机号码")]
        public string Phone { get; set; }

        [TitleName("出生日期")]
        public DateTime? Birthday { get; set; }

        [TitleName("身高(米)")]
        public decimal Height { get; set; }

        [TitleName("登记日期")]
        public DateTime CreateTime { get; set; }
    }
}
