using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOIHelper
{
    public class DateTimeDataType : IDataType
    {
        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="columnType"></param>
        public DateTimeDataType(ColumnType columnType) : base(columnType)
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
            cell.CellStyle = style;
            var _dateTemp = DateTime.Today;
            if (DateTime.TryParse(value, out _dateTemp))
            {
                cell.SetCellValue(_dateTemp);
                cell.CellStyle = style;
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
        public override ICellStyle CreateCellStyle(HSSFWorkbook workBook)
        {
            ICellStyle style = workBook.CreateCellStyle();

            IDataFormat datastyle = workBook.CreateDataFormat();
            style.DataFormat = datastyle.GetFormat("yyyy/MM/dd hh:mm:ss");
            return style;
        }
    }
}
