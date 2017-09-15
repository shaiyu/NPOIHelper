using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    public enum ColumnType
    {
        /// <summary>
        /// 默认，不作任何转换
        /// </summary>
        Default = 0,
        /// <summary>
        /// 默认数字格式
        /// </summary>
        Number = 1,
        /// <summary>
        /// 2 位小数 --> 未实现，采用默认数字格式
        /// </summary>
        NumDecimal2 = 2,
        /// <summary>
        /// 百分比
        /// </summary>
        NumberPercentage = 3,
        /// <summary>
        /// 科学计数法
        /// </summary>
        NumberScientificNotation = 4,
        /// <summary>
        /// 日期格式
        /// </summary>
        Date = 10,
        /// <summary>
        /// 日期时间格式
        /// </summary>
        DateTime = 11,
        DateFile = 12,
        /// <summary>
        /// 字符串，暂不作任何转换
        /// </summary>
        String = 20
    }
}
