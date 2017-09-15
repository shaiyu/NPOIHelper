using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    /// <summary>
    /// Model 约束
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnTypeAttribute : Attribute
    {
        /// <summary>
        /// 构造
        /// </summary>
        public ColumnTypeAttribute()
        {
            this.Hide = false; 
            this.Type = ColumnType.Default;
        }

        /// <summary>
        /// 导出名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 导出类型
        /// </summary>
        public ColumnType Type { get; set; }
        /// <summary>
        /// 是否隐藏 默认不隐藏,导出
        /// </summary>
        public bool Hide { get; set; }

    }
}
