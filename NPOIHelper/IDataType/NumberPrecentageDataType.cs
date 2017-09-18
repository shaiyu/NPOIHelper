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
    /// 百分比样式
    /// </summary>
    public class NumberPrecentageDataType : IDataType
    {

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="columnType"></param>
        public NumberPrecentageDataType(ColumnType columnType) : base(columnType)
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

            var _numberPTemp = 0D;
            if (Double.TryParse(value.Replace("%", ""), out _numberPTemp))
            {
                cell.SetCellValue(value.Contains("%") ? _numberPTemp / 100 : _numberPTemp);
            }
            else
            {
                cell.SetCellValue(value);
            }
            cell.CellStyle = style;
        }

        /// <summary>
        /// 获取对应Data的样式
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public override ICellStyle CreateCellStyle(HSSFWorkbook workBook)
        {
            ICellStyle style = workBook.CreateCellStyle();

            IDataFormat datastyle = workBook.CreateDataFormat();
            style.DataFormat = datastyle.GetFormat("0.00%");
            return style;
        }
    }
}
