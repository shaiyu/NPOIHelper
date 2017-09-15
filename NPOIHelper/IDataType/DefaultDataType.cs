using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOIHelper
{
    public class DefaultDataType : IDataType
    {

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="columnType"></param>
        public DefaultDataType(ColumnType columnType) : base(columnType)
        {
        }

        /// <summary>
        /// 为指定单元格填充值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="style"></param>
        public override void SetCellValue(ICell cell, string value, ICellStyle style)
        {
            cell.SetCellValue(value);
        }

        /// <summary>
        /// 获取对应Data的样式
        /// 默认不启用样式设置
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public override ICellStyle CreateCellStyle(HSSFWorkbook workBook)
        {
            return null;
        }
    }
}
