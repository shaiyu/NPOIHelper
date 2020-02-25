using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    public class XSSFExcelHelper : ExcelHelper
    {
        XSSFWorkbook xssfWorkbook;

        /// <summary>
        /// 添加Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public override void Add(ISheet sheet)
        {
            xssfWorkbook.Add(sheet);
        }

        /// <summary>
        /// 创建并初始化工作簿
        /// </summary>
        public override void Init()
        {
            xssfWorkbook = new XSSFWorkbook();
            WorkBook = xssfWorkbook;

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
            res.AddHeader("Content-Disposition", "attachment; filename=" + this.ExcelName + ".xls");
            res.ContentEncoding = System.Text.Encoding.GetEncoding("GBK");
            res.ContentType = "application/vnd.ms-excel;charset=GBK";

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
                var properties =  xssfWorkbook.GetProperties();
                var coreProperties = properties.CoreProperties;
                var customProperties = properties.CustomProperties;
                //var extentProperties = properties.ExtendedProperties;

                //文档摘要信息
                if (!customProperties.Contains("Company"))
                {
                    customProperties.AddProperty("Company", "畅途汽车技术服务有限公司");
                }

                coreProperties.Creator = "畅途"; //填加xlsx文件作者信息
                coreProperties.Keywords = "Labbor"; //填加xls文件最后保存者信息
                coreProperties.Title = "畅途汽车技术服务有限公司"; //填加xls文件标题信息
                coreProperties.Subject = "畅途汽车技术服务有限公司";//填加文件主题信息
                coreProperties.Created = DateTime.Now;
            }
            #endregion
        }

    }
}
