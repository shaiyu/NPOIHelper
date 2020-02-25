using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    /// <summary>
    /// <para>added by Labbor on 20170613 sheet表的抽象类</para>
    /// </summary>
    public abstract class Sheet
    {
        /// <summary>
        /// 构造
        /// </summary>
        public Sheet(IWorkbook workBook, string sheetName)
        {
            this.WorkBook = workBook;
            this.SheetName = this.InitialSheetName = sheetName;
            //默认导出标题
            this.ShowTitle = true;
            //默认冻结首行
            this.FreezePane = true;
        }

        /// <summary>
        /// Excel Sheet的最大行数 2003版之前最大是2^16 = 65535
        /// </summary>
        public static int MaxRow = 65535;
        //public static int MaxRow = 20;

        /// <summary>
        /// 当前Excel文件实例
        /// </summary>
        public IWorkbook WorkBook { get; set; }

        /// <summary>
        /// 当前表实例
        /// </summary>
        public ISheet SheetThis { get; set; }

        /// <summary>
        /// 当前表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 初始表名称
        /// </summary>
        public string InitialSheetName { get; set; }

        /// <summary>
        /// 是否创建表头
        /// </summary>
        public bool ShowTitle { get; set; }

        /// <summary>
        /// 是否冻结首行 冻结哪行？哪列？
        /// </summary>
        public bool FreezePane { get; set; }

        /// <summary>
        /// 对应需要导出的每一列
        /// </summary>
        public Column[] Columns { get; set; }

        /// <summary>
        /// 生成带有数据的Sheet
        /// </summary>
        public void Build()
        {
            CheckAndReBuildColumns();
            CreateSheet();
            FillData();
        }

        /// <summary>
        /// 创建表
        /// </summary>
        public void CreateSheet()
        {
            if (string.IsNullOrWhiteSpace(this.SheetName))
            {
                this.SheetName = "默认名称";
            }
            CheckAndReBuildSheetName();
            SheetThis = WorkBook.CreateSheet(SheetName);
            SheetThis.DefaultColumnWidth = 20;
            SheetThis.DefaultRowHeightInPoints = 20;
            if (ShowTitle)
            {
                //调用创建标题的方法
                CreateTitle();
                if (FreezePane)
                {
                    //冻结首行
                    SheetThis.CreateFreezePane(0, 1);
                }
            }
        }

        /// <summary>
        /// 检查重设表名
        /// </summary>
        public void CheckAndReBuildSheetName()
        {
            if (this.WorkBook.GetSheet(this.SheetName) != null)
            {
                this.SheetName += ">";
                this.CheckAndReBuildSheetName();
            }
        }

        /// <summary>
        /// 如果指定了列, 则检查数据源是否存在该列
        /// </summary>
        public abstract void CheckAndReBuildColumns();

        /// <summary>
        /// 创建标题
        /// </summary>
        public virtual void CreateTitle()
        {
            IRow row = SheetThis.CreateRow(0);
            var i = 0;
            if (Columns == null || Columns.Length == 0)
            {
                this.InitDefaultColumns();
            }
            if (Columns != null)
            {
                foreach (var col in Columns)
                {
                    ICell cell = row.CreateCell(i++);
                    cell.SetCellValue(col.ColTitleName);
                    cell.CellStyle = GetTitleStyle(WorkBook);
                }
            }
        }

        internal abstract void InitDefaultColumns();

        /// <summary>
        /// 填充数据
        /// </summary>
        public void FillData()
        {
            BeforeFill();
            Fill();
            AfterFill();
        }

        /// <summary>
        /// 数据填充前
        /// </summary>
        public virtual void BeforeFill()
        {

        }

        /// <summary>
        /// 填充数据
        /// </summary>
        public abstract void Fill();

        /// <summary>
        /// 数据填充后
        /// </summary>
        public virtual void AfterFill()
        {
            //没有此句，则不会刷新出公式计算结果
            this.SheetThis.ForceFormulaRecalculation = true;
        }

        /// <summary>
        /// 设置标题Cell样式
        /// </summary>
        /// <param name="wk"></param>
        /// <returns></returns>
        public static ICellStyle GetTitleStyle(IWorkbook wk)
        {
            IFont font = wk.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 12;
            font.Boldweight = short.MaxValue;
            ICellStyle style = wk.CreateCellStyle();
            style.SetFont(font);
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            return style;
        }

        /// <summary>
        /// 类型字典
        /// </summary>
        private Dictionary<ColumnType, ICellStyle> dicCellStyle = new Dictionary<ColumnType, ICellStyle>();

        /// <summary>
        /// 获取ICellStyle 有则直接取出,无则存入字典
        /// </summary>
        /// <param name="dataType"></param>
        /// <returns></returns>
        public ICellStyle GetCellStyle(IDataType dataType)
        {
            ICellStyle cellStyle = null;
            if (dicCellStyle.ContainsKey(dataType.ColumnType))
            {
                cellStyle = dicCellStyle[dataType.ColumnType];
            }
            else
            {
                cellStyle = dataType.CreateCellStyle(this.WorkBook);
                dicCellStyle.Add(dataType.ColumnType, cellStyle);
            }
            return cellStyle;
        }

    }
}
