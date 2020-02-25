using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    public class HSSFExcelHelper : ExcelHelper
    {
        HSSFWorkbook hssfWorkbook;

        /// <summary>
        /// 添加Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public override void Add(ISheet sheet)
        {
            hssfWorkbook.Add(sheet);
        }

        /// <summary>
        /// 创建并初始化工作簿
        /// </summary>
        public override void Init()
        {
            hssfWorkbook = new HSSFWorkbook();
            WorkBook = hssfWorkbook;

            IFont font = WorkBook.GetFontAt(0);
            font.FontName = "宋体";

            SetInformation();
        }

        public override void PreReport()
        {
        }

        public override void Report()
        {
            PreReport();

            //下载报表
            var res = System.Web.HttpContext.Current.Response;
            res.Clear();
            res.Buffer = true;
            res.Charset = "GBK";
            res.AddHeader("Content-Disposition", "attachment; filename=" + this.ExcelName + ".xlsx");
            res.ContentEncoding = System.Text.Encoding.GetEncoding("GBK");
            res.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=GBK";

            //Wirte Stream 
            this.WorkBook.Write(res.OutputStream);

            res.Flush();
            res.End();
        }

        /// <summary>
        /// <para>配置文件信息</para>
        /// <para>si.Author 填加xls文件作者信息</para>
        /// <para>si.LastAuthor 填加xls文件最后保存者信息</para>
        /// <para>si.Comments </para>
        /// <para>si.Title 填加xls文件标题信息</para>
        /// <para>si.Subject 填加文件主题信息</para>
        /// </summary>
        public override void SetInformation()
        {
            #region 附加信息
            {
                //文档摘要信息
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "畅途汽车技术服务有限公司";
                hssfWorkbook.DocumentSummaryInformation = dsi;
                //
                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "畅途"; //填加xls文件作者信息
                si.ApplicationName = "NPOI程序"; //填加xls文件创建程序信息
                si.LastAuthor = "Labbor"; //填加xls文件最后保存者信息
                si.Comments = "畅途汽车技术服务有限公司所有";
                si.Title = "畅途汽车技术服务有限公司"; //填加xls文件标题信息
                si.Subject = "畅途汽车技术服务有限公司";//填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                hssfWorkbook.SummaryInformation = si;
            }
            #endregion
        }

    }
}
