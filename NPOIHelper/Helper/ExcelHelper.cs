using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
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
    public class ExcelHelper : IHelper
    {
        HSSFWorkbook workbook;

        /// <summary>
        /// 当前工作表
        /// </summary>
        public HSSFWorkbook WorkBook { get { return this.workbook; } set { this.workbook = value; } }

        /// <summary>
        /// Excel表名
        /// </summary>
        public string ExcelName { get; set; }

        /// <summary>
        /// 构造方法
        /// </summary>
        public ExcelHelper()
        {
            this.ExcelName = "默认名称";
            this.Init();
        }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="_ExcelName"></param>
        public ExcelHelper(string _ExcelName)
        {
            this.ExcelName = _ExcelName;
            this.Init();
        }

        /// <summary>
        /// 添加Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public void Add(ISheet sheet)
        {
            workbook.Add(sheet);
        }

        /// <summary>
        /// 添加Sheet  MaxLine = 65535
        /// </summary>
        /// <param name="_SheetName"></param>
        /// <param name="data"></param>
        /// <param name="columns"></param>
        public void Add(string _SheetName, DataTable data, Column[] columns = null)
        {
            int len = GetLen(data == null ? 0 : data.Rows.Count);
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
            int len = GetLen(data == null ? 0 : data.Count);
            for (int i = 0; i < len; i++)
            {
                var sheetName = _SheetName + (i > 0 ? i + "" : "");
                var sheet = new ListSheet<T>(workbook, data.Skip(Sheet.MaxRow * i).Take(Sheet.MaxRow).ToList(), _SheetName);
                sheet.Columns = columns;
                sheet.Build();
            }
        }

        /// <summary>
        /// 创建并初始化工作簿
        /// </summary>
        private void Init()
        {
            workbook = new HSSFWorkbook();
            IFont font = workbook.GetFontAt((short)0);
            font.FontName = "宋体";
        }

        /// <summary>
        /// <para>Http导出 清空响应流,并写入数据</para>
        /// </summary>
        public void Report()
        {
            //下载报表
            var res = System.Web.HttpContext.Current.Response;
            res.Clear();
            res.Buffer = true;
            res.Charset = "GBK";
            res.AddHeader("Content-Disposition", "attachment; filename=" + this.ExcelName + DateTime.Now.ToShortDateString() + ".xls");
            res.ContentEncoding = System.Text.Encoding.GetEncoding("GBK");
            res.ContentType = "application/ms-excel;charset=GBK";

            //Wirte Stream 
            this.WorkBook.Write(res.OutputStream);

            res.Flush();
            res.End();
        }

        /// <summary>
        /// <para>WinForm等客户端导出</para>
        /// </summary>
        /// <param name="fileName"></param>
        public void ReportClient(string fileName)
        {

            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                this.WorkBook.Write(fs);
                fs.Flush();
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
        /// 配置文件信息
        /// </summary>
        public void SetInformation()
        {
            #region 附加信息
            {
                //文档摘要信息
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "畅途汽车技术服务有限公司";
                workbook.DocumentSummaryInformation = dsi;
                //
                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "畅途"; //填加xls文件作者信息
                si.ApplicationName = "NPOI程序"; //填加xls文件创建程序信息
                si.LastAuthor = "Labbor"; //填加xls文件最后保存者信息
                si.Comments = "畅途汽车技术服务有限公司所有"; //填加xls文件作者信息
                si.Title = "畅途汽车技术服务有限公司"; //填加xls文件标题信息
                si.Subject = "畅途汽车技术服务有限公司";//填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }
            #endregion
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
