using NPOI.SS.UserModel;
using NPOIHelper.Enums;

namespace NPOIHelper
{

    /// <summary>
    /// added by Labbor on 20170602
    /// </summary>
    public interface INPOIHelper
    {
        string FileName { get; set; }
        string Extension { get; }
        string FullName { get; }
        string ContentType { get; }
        IWorkbook WorkBook { get; }
        NPOIType Type { get; }

        /// <summary>
        /// to byte array
        /// </summary>
        byte[] ToArray();
        
        /// <summary>
        /// Client导出报表
        /// </summary>
        void ReportClient(string fileName);

        /// <summary>
        /// 设置报表信息
        /// </summary>
        void SetInformation();
    }
}
