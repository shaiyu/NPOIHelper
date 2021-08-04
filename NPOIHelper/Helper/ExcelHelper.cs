using NPOI;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOIHelper.Enums;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    /// <summary>
    /// 帮助类
    /// </summary>
    internal abstract class ExcelHelper : IExcelHelper
    {
        private IWorkbook workbook;

        /// <summary>
        /// 当前工作表
        /// </summary>
        public IWorkbook WorkBook { get { return this.workbook; } set { this.workbook = value; } }

        /// <summary>
        /// Excel表名
        /// </summary>
        public string FileName { get; set; }
        public string Extension => Type.ToString();
        public string FullName
        {
            get
            {
                return $"{FileName}.{Extension}";
            }
        }
        public string ContentType => Type.GetDefaultValue<string>();
        public NPOIType Type { get; set; }

        /// <summary>
        /// 构造方法
        /// </summary>
        public ExcelHelper()
        {
            FileName = "默认名称";
            Init();
        }

        /// <summary>
        /// 创建并初始化工作簿
        /// </summary>
        public abstract void Init();

        public abstract void SetInformation();

        /// <summary>
        /// 添加Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public abstract void Add(ISheet sheet);

        /// <summary>
        /// 添加Sheet  MaxLine = 65535
        /// </summary>
        /// <param name="_SheetName"></param>
        /// <param name="data"></param>
        /// <param name="columns"></param>
        public void Add(string _SheetName, DataTable data, Column[] columns = null)
        {
            data = data == null ? new DataTable() : data;
            int len = GetLen(data.Rows.Count);
            for (int i = 0; i < len; i++)
            {
                var sheetName = _SheetName + (i > 0 ? i + "" : "");
                var sheet = new DataTableSheet(workbook, SplitDataTable(data, Sheet.MaxRow * i, Sheet.MaxRow), sheetName);
                sheet.Columns = columns;
                sheet.Build();
            }
        }

        /// <summary>
        /// 添加Sheet  MaxLine = 65535
        /// </summary>
        /// <param name="_SheetName"></param>
        /// <param name="data"></param>
        /// <param name="columns"></param>
        public void Add<T>(string _SheetName, List<T> data, Column[] columns = null)
        {
            data = data == null ? new List<T>() : data;
            int len = GetLen(data.Count);
            for (int i = 0; i < len; i++)
            {
                var sheetName = _SheetName + (i > 0 ? i + "" : "");
                var sheet = new ListSheet<T>(workbook, data.Skip(Sheet.MaxRow * i).Take(Sheet.MaxRow).ToList(), sheetName);
                sheet.Columns = columns;
                sheet.Build();
            }
        }


        /// <summary>
        /// 
        /// </summary>
        public void PreReport()
        {

        }

        /// <summary>
        /// <para>Http导出 清空响应流,并写入数据</para>
        /// </summary>
        public byte[] ToArray()
        {
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// <para>WinForm等客户端导出</para>
        /// </summary>
        /// <param name="fileName"></param>
        public void ReportClient(string fileName)
        {
            PreReport();

            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                this.WorkBook.Write(fs);
                if (fs.CanRead || fs.CanSeek || fs.CanWrite)
                {
                    fs.Flush();
                    fs.Close();
                }
            }

            ////下载报表
            //using (MemoryStream ms = new MemoryStream()) {
            //    this.WorkBook.Write(ms);
            //    using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            //    {
            //        byte[] data = ms.ToArray();
            //        fs.Write(data, 0, data.Length);
            //        fs.Flush();
            //    }
            //    ms.Flush();
            //    ms.Position = 0;
            //}
        }

        /// <summary>
        /// 获取分页数 因每个Excel Sheet最多有65535行
        /// </summary>
        /// <param name="rowsCount"></param>
        /// <returns></returns>
        private int GetLen(int rowsCount)
        {
            int len = (int)Math.Ceiling((double)rowsCount / Sheet.MaxRow);
            return len < 1 ? 1 : len;
        }

        private DataTable SplitDataTable(DataTable dt, int start, int len)
        {
            if (dt == null || len <= 0)
                return null;
            int max = dt.Rows.Count;
            if (start > max)
                return null;
            start = start < 0 ? 0 : start;
            if (start + len > max)
                len = max - start;
            var newDt = dt.Clone();
            dt.Rows.Cast<DataRow>().Skip(start).Take(len).ToList().ForEach(row => newDt.ImportRow(row));
            return newDt;
        }
    }



}
