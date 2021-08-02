using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOIHelper
{
   internal class NPOITypeUtils
    {
        public static NPOIType GetType(string fileName)
        {
            var ext = Path.GetExtension(fileName);
            return GetTypeByExt(ext);
        }

        public static NPOIType GetTypeByContentType(string contentType)
        {
            var ext = MimeTypeMap.GetExtension(contentType);
            return GetTypeByExt(ext);
        }

        private static NPOIType GetTypeByExt(string ext)
        {
            if (ext == "." + NPOIType.xlsx.ToString())
            {
                return NPOIType.xlsx;
            }
            else if (ext == "." + NPOIType.xls.ToString())
            {
                return NPOIType.xls;
            }
            throw new NotSupportedException("不支持的Excel文件类型");
        }
    }
}
