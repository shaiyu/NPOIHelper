using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    /// <summary>
    /// 数据类型
    /// </summary>
    public abstract class IDataType
    {

        /// <summary>
        /// 类型名
        /// </summary>
        public ColumnType ColumnType { get; set; }

        public IDataType(ColumnType type)
        {
            ColumnType = type;
        }

        /// <summary>
        /// 设置表格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="style"></param>
        public abstract void SetCellValue(ICell cell,string value,ICellStyle style);

        /// <summary>
        /// 构建属于当前类型的CellType
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public abstract ICellStyle CreateCellStyle(HSSFWorkbook workBook);

    }
}
