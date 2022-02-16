using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelperTest
{
    public class ImportUser2
    {
        [TitleName("ID")]
        public int Id { get; set; }

        [TitleName("姓名")]
        public string Name { get; set; }

        [TitleName("年龄")]
        public int Age { get; set; }
    }
}
