using NPOI.SS.UserModel;

namespace NPOIHelper
{

    /// <summary>
    /// added by Labbor on 20170602
    /// </summary>
    public interface INPOIHelper
    {
        /// <summary>
        /// 文件名
        /// </summary>
        string FileName { get; set; }
        /// <summary>
        /// 扩展名
        /// </summary>
        string Extension { get; }
        /// <summary>
        /// 包含扩展名的文件名称
        /// </summary>
        string FullName { get; }

        /// <summary>
        /// 文件的Http ContentType
        /// </summary>
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
