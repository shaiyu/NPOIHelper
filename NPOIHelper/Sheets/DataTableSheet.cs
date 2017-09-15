using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    /// <summary>
    /// <para>added by Labbor on 20170614 DataTable Sheet类</para>
    /// </summary>
    public class DataTableSheet : Sheet
    {
        /// <summary>
        /// 数据源
        /// </summary>
        public DataTable Data { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public DataTableSheet(HSSFWorkbook workBook, DataTable data, string sheetName) : base(workBook, sheetName)
        {
            this.Data = data;
        }

        /// <summary>
        /// 构造
        /// </summary>
        public DataTableSheet(HSSFWorkbook workBook, DataTable data, string sheetName, Column[] columns) : base(workBook, sheetName)
        {
            this.Data = data;
            this.Columns = columns;
        }

        /// <summary>
        /// 如果指定了列, 则检查是否存在该列
        /// </summary>
        public override void CheckAndReBuildColumns()
        {
            if (this.Columns==null)
                return;
            if (Data != null && Data.Columns.Count > 0)
            {
                List<Column> columnList = new List<Column>(1);
                for (int i = 0; i < this.Columns.Count(); i++)
                {
                    if (Data.Columns.Cast<DataColumn>().Where(s=>s.ColumnName == this.Columns[i].ColName).FirstOrDefault() != null)
                    {
                        columnList.Add(this.Columns[i]);
                    }
                }
                this.Columns = columnList.ToArray();
            }
        }

        /// <summary>
        /// 如果不指定列,则获取Table列
        /// </summary>
        internal override void InitDefaultColumns()
        {
            if (Data != null && Data.Columns.Count > 0)
            {
                this.Columns = new Column[Data.Columns.Count];
                for (int i = 0; i < Data.Columns.Count; i++)
                {
                    this.Columns[i] = new Column(Data.Columns[i].ColumnName, Data.Columns[i].ColumnName, ColumnType.Default);
                }
            }
        }

        /// <summary>
        /// 填充数据
        /// </summary>
        public override void Fill()
        {
            if (Data != null && Data.Rows.Count > 0)
            {
                //导入初始行  0 位第一行标题
                var start = 1;
                // 表格值
                string colsValue = null;
                ICellStyle cellStyle = null;
                foreach (DataRow dr in Data.Rows)
                {
                    IRow _row = SheetThis.CreateRow(start);
                    for (int i = 0; i < this.Columns.Length; i++)
                    {
                        var cell = _row.CreateCell(i);
                        // get cell value
                        colsValue = dr[this.Columns[i].ColName] + "";

                        // format cell value
                        this.Columns[i].FormattedValue(cell, dr, start, ref colsValue);

                        // 获取单元格样式 默认从字典中读取
                        cellStyle = GetCellStyle(this.Columns[i].DataType);
                        // 设置单元格数据
                        this.Columns[i].DataType.SetCellValue(cell, colsValue, cellStyle);
                    }
                    start++;
                }
            }
        }

    }
}
