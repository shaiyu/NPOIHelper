using NPOIHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AspNetCoreSample.Models
{
    public class ImportDto
    {
        [TitleName("推荐类型")]
        public string Type { get; set; }

        [TitleName("推荐机构/课程id")]
        public string IdsString { get; set; }

        [TitleName("投放学校id")]
        public Guid? SId { get; set; }

        [TitleName("投放学部id")]
        public string ExtIdsString { get; set; }

        [TitleName("开始时间")]
        public DateTime StartTime { get; set; }

        [TitleName("结束时间")]
        public DateTime EndTime { get; set; }
    }
}
