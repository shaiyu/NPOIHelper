using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelperTest.Model
{
    public class ImportRecommend
    {
        [TitleName("开始时间")]
        public DateTime StartTime { get; set; }

        [TitleName("结束时间")]
        public DateTime EndTime { get; set; }

    }
}
