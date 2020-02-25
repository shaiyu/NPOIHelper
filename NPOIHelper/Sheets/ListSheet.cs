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
    /// <para>added by Labbor on 20170614 ListData Sheet类</para>
    /// </summary>
    public class ListSheet<T> : Sheet
    {
        /// <summary>
        /// 获取 T 的类型
        /// </summary>
        /// <returns></returns>
        public Type GetClazz()
        {
            var clazz = typeof(T);
            if (clazz.Name.ToLower() == "object" && Data.Count > 0)
            {
                clazz = Data.FirstOrDefault().GetType();
            }
            return clazz;
        }

        /// <summary>
        /// 数据源
        /// </summary>
        public List<T> Data { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ListSheet(IWorkbook workBook, List<T> data, string sheetName) : base(workBook, sheetName)
        {
            this.Data = data;
        }

        /// <summary>
        /// 如果指定了列, 则检查是否存在该列
        /// </summary>
        public override void CheckAndReBuildColumns()
        {
            if (this.Columns == null)
                return;
            var clazz = GetClazz();
            List<Column> columnList = new List<Column>(1);
            for (int i = 0; i < this.Columns.Length; i++)
            {
                //当前字段在列表属性内
                if (clazz.GetProperty(this.Columns[i].ColName) != null)
                {
                    columnList.Add(this.Columns[i]);
                }
            }
            this.Columns = columnList.ToArray();
        }

        /// <summary>
        /// 重写初始化列的方法 如果已指定列, 则不会调用
        /// </summary>
        internal override void InitDefaultColumns()
        {
            if (Data != null && Data.Count > 0)
            {
                Type clazz = GetClazz();
                Type clazz2 = typeof(ColumnTypeAttribute);
                string titleName = "";
                List<Column> columnList = new List<Column>(1);
                ColumnType type = ColumnType.Default;
                System.Reflection.PropertyInfo[] properties = clazz.GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    var p = properties[i];
                    var attr = (ColumnTypeAttribute)p.GetCustomAttributes(clazz2, false).FirstOrDefault();
                    if (attr == null)
                    {
                        titleName = p.Name;
                        type = ColumnType.Default;
                        columnList.Add(new Column(p.Name, titleName, type));
                    }
                    else if (!attr.Hide)
                    {
                        titleName = attr.Name == null ? p.Name : attr.Name;
                        type = attr.Type;
                        columnList.Add(new Column(p.Name, titleName, type));
                    }
                }
                this.Columns = columnList.ToArray();
            }
        }

        /// <summary>
        /// 填充数据
        /// </summary>
        public override void Fill()
        {
            if (Data != null && Data.Count > 0)
            {
                // 导入初始行  0 位第一行标题
                var start = 1;
                // 表格值
                string colsValue = null;
                ICellStyle cellStyle = null;
                var clazz = GetClazz();
                foreach (T t in Data)
                {
                    IRow _row = SheetThis.CreateRow(start);
                    for (int i = 0; i < this.Columns.Length; i++)
                    {
                        var cell = _row.CreateCell(i);
                        // get cell value
                        try
                        {
                            colsValue = clazz.GetProperty(this.Columns[i].ColName).GetValue(t) + "";
                        }
                        catch (Exception)
                        {
                            //未找到该属性
                        }

                        // format cell value
                        this.Columns[i].FormattedValue(cell, t, start, ref colsValue);

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
