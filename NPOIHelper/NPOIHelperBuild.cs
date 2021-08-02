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
    public class NPOIHelperBuild
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
        /// 读取excel到IEnumerable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelFileName"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static IEnumerable<T> ReadExcel<T>(string excelFileName, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetType(excelFileName);
            return new ListExcelReader(excelFileName, type, columnLength).Read<T>();
        }

        /// <summary>
        /// 读取excel到IEnumerable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static IEnumerable<T> ReadExcel<T>(Stream stream, string contentType, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetTypeByContentType(contentType);
            return new ListExcelReader(stream, type, columnLength).Read<T>();
        }


        /// <summary>
        /// 读取Excel到DataTable
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static DataTable ReadExcel(string excelFileName, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetType(excelFileName);
            return new DataTableExcelReader(excelFileName, type, columnLength).Read();
        }


        /// <summary>
        /// 读取excel到IEnumerable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static DataTable ReadExcel(Stream stream, string contentType, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetTypeByContentType(contentType);
            return new DataTableExcelReader(stream, type, columnLength).Read();
        }

    }
}
