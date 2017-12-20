using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    /// <summary>
    /// added by Labbor on 20170602  Excel导出的列
    /// </summary>
    public class Column
    {
        /// <summary>
        /// 列字段 --> state
        /// </summary>
        public string ColName { get; set; }
        /// <summary>
        /// 列标题名称 --> 状态
        /// </summary>
        public string ColTitleName { get; set; }
        /// <summary>
        /// 列类型 --> 字段类型  使用枚举,仅为使用方便
        /// </summary>
        public ColumnType ColType { get; set; }
        /// <summary>
        /// 具体使用的数据类型
        /// </summary>
        public IDataType DataType { get; set; }

        /// <summary>
        /// 列数据源 --> 自动取值
        /// </summary>
        public Dictionary<string, string> ColDataSourse { get; set; }

        /// <summary>
        /// 默认值 暂未使用
        /// </summary>
        public string DefaultValue { get; set; }

        /// <summary>
        /// 使用定值 暂未使用
        /// </summary>
        public string FixedValue { get; set; }

        /// <summary>
        /// 转换函数 DataRow ColsValue Index
        /// </summary>
        public Func<object, int, string> Func { get; set; }

        /// <summary>
        /// 值是否是公式
        /// </summary>
        public bool IsFormula { get; set; }

        private Column() { }

        /// <summary>
        /// 基本构造
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="colTitleName"></param>
        /// <param name="colType"></param>
        public Column(string colName, string colTitleName, ColumnType colType = ColumnType.Default)
        {
            this.ColName = colName;
            this.ColTitleName = colTitleName;
            this.ColType = colType;
            this.DataType = TypeFactory.Get(colType);
        }

        /// <summary>
        /// 基本构造
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="colTitleName"></param>
        /// <param name="colType"></param>
        /// <param name="colDataSourse"></param>
        public Column(string colName, string colTitleName, ColumnType colType, Dictionary<string, string> colDataSourse)
            : this(colName, colTitleName, colType)
        {
            this.ColDataSourse = colDataSourse;
        }

        /// <summary>
        /// 基本构造
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="colTitleName"></param>
        /// <param name="colType"></param>
        /// <param name="func">自定义转换函数 object -> Entity</param>
        [Obsolete]
        public Column(string colName, string colTitleName, ColumnType colType, Func<object, string> func)
            : this(colName, colTitleName, colType)
        {
            this.Func = (t, index) => {
                return func.Invoke(t);
            };
        }

        /// <summary>
        /// 基本构造
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="colTitleName"></param>
        /// <param name="colType"></param>
        /// <param name="func">自定义转换函数 object -> Entity  int -> ExcelRowIndex, 加上了标题行</param>
        public Column(string colName, string colTitleName, ColumnType colType, Func<object, int, string> func)
            : this(colName, colTitleName, colType)
        {
            this.Func = func;
        }

        /// <summary>
        /// 基本构造
        /// </summary>
        /// <param name="colName"></param>
        /// <param name="colTitleName"></param>
        /// <param name="func">自定义转换函数</param>
        public Column(string colName, string colTitleName, Func<object, int, string> func)
            : this(colName, colTitleName)
        {
            this.Func = func;
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="rowData">行数据</param>
        /// <param name="rowIndex">行号</param>
        /// <param name="colsValue">当前值</param>
        public void FormattedValue(ICell cell, object rowData, int rowIndex, ref string colsValue)
        {
            rowIndex += 1; // 实际行号 标题行+1
            // format cell value
            if (this.Func != null)
            {
                //transform data
                colsValue = this.Func(rowData, rowIndex);
            }
            else
            {
                //如果有匹配的数据源，则更换数据    
                if (ColDataSourse != null && ColDataSourse.ContainsKey(colsValue))
                {
                    colsValue = ColDataSourse[colsValue];
                }
            }
            //设置公式
            //值是公式 则使用公式
            if (IsFormula)
            {
               //cell.SetCellFormula(colsValue);
               cell.SetCellFormula(string.Format(colsValue, rowIndex));
            }
            //return colsValue;
        }

    }
}
