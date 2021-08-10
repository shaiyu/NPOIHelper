using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace NPOIHelper
{
    public partial class NPOIHelperBuild
    {
        /// <summary>
        /// 获取Excel导出帮助类
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static IExcelHelper GetHelper(NPOIType type = NPOIType.xlsx)
        {
            if (type == NPOIType.xlsx)
            {
                return new XSSFExcelHelper() { Type = type };
            }
            return new HSSFExcelHelper() { Type = type };
        }

        /// <summary>
        /// 获取Excel导出帮助类
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static IExcelHelper GetHelper(string fileName, NPOIType type = NPOIType.xlsx)
        {
            if (type == NPOIType.xlsx)
            {
                return new XSSFExcelHelper() { Type = type, FileName = fileName };
            }
            return new HSSFExcelHelper() { Type = type, FileName = fileName };
        }

        /// <summary>
        /// 获取Excel读取帮助类
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <returns></returns>
        public static ExcelReader GetReader(string excelFileName)
        {
            var type = NPOITypeUtils.GetType(excelFileName);
            return new ExcelReader(excelFileName, type);
        }

        /// <summary>
        /// 获取Excel读取帮助类
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <returns></returns>
        public static ExcelReader GetReader(Stream stream, string contentType)
        {
            var type = NPOITypeUtils.GetTypeByContentType(contentType);
            return new ExcelReader(stream, type);
        }
    }
}
