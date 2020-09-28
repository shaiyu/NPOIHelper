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
    /// 保留两位小数, 不可用, 启用了类型缓存, 所以只会实例化第一个数字类型
    /// </summary>
    public class NumberNDataType : IDataType
    {
        private int Digital { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="columnType"></param>
        /// <param name="digital">小数点后的位数</param>
        public NumberNDataType(ColumnType columnType, int digital) : base(columnType)
        {
            this.Digital = digital;
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
            cell.CellStyle = style;
        }

        /// <summary>
        /// 获取对应Data的样式
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public override ICellStyle CreateCellStyle(IWorkbook workBook)
        {
            ICellStyle style = workBook.CreateCellStyle();
            IDataFormat datastyle = workBook.CreateDataFormat();
            var format = getFormat();
            style.DataFormat = datastyle.GetFormat(format);
            return style;
        }

        private string getFormat()
        {
            var format = "0";
            if (Digital > 0)
            {
                format += ".";
                for (int i = 0; i < Digital; i++)
                {
                    format += "0";
                }
            }
            return format;
        }
    }
}
