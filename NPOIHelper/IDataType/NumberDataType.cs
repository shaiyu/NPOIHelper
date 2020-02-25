using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOIHelper
{
    /// <summary>
    /// 普通数字样式
    /// </summary>
    public class NumberDataType : IDataType
    {
        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="columnType"></param>
        public NumberDataType(ColumnType columnType) : base(columnType)
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
            var _numberTemp = 0D;
            if (Double.TryParse(value, out _numberTemp))
            {
                cell.SetCellValue(_numberTemp);
            }
            else
            {
                cell.SetCellValue(value);
            }
        }

        /// <summary>
        /// 获取对应Data的样式
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public override ICellStyle CreateCellStyle(IWorkbook workBook)
        {
            return null;
        }
    }
}
