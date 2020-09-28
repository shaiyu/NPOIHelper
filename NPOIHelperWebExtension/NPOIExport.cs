using NPOIHelper;
using System;

namespace NPOIHelperWebExtension
{

#if NET45
    using System.Web;
    public class NPOIExport
    {
        public static void Export(HttpResponse response, IExcelHelper helper)
        {
            //下载报表
            var res = response;
            res.Clear();
            res.Buffer = true;
            res.Charset = "utf8";
            res.AddHeader("Content-Disposition", $"attachment; filename={helper.FileName}.{helper.FullName}");
            res.ContentEncoding = System.Text.Encoding.GetEncoding("utf8");
            res.ContentType = helper.ContentType;

            //Wirte Stream 
            helper.WorkBook.Write(res.OutputStream);

            res.Flush();
            res.End();
        }

    }
#else
    using Microsoft.AspNetCore.Http;
    public class NPOIExport
    {
        public static void Export(HttpResponse response, IExcelHelper helper)
        {
            //下载报表
            var res = response;

            //res.Body.Flush();
            res.ContentType = helper.ContentType;
            res.Headers.Add("Content-Disposition", $"attachment; filename={helper.FullName};utf8");

            var bytes = helper.ToArray();
            res.Body.WriteAsync(bytes, 0, bytes.Length);
        }
    }
#endif
}
