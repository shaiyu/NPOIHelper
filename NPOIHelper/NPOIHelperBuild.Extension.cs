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
        /// 读取excel到IEnumerable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelFileName"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static IEnumerable<T> ReadExcel<T>(string excelFileName, int sheetIndex = 0, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetType(excelFileName);
            return new ExcelReader(excelFileName, type).ReadSheet<T>(sheetIndex, columnLength);
        }

        /// <summary>
        /// 读取excel到IEnumerable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static IEnumerable<T> ReadExcel<T>(Stream stream, string contentType, int sheetIndex = 0, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetTypeByContentType(contentType);
            return new ExcelReader(stream, type).ReadSheet<T>(sheetIndex, columnLength);
        }


        /// <summary>
        /// 读取Excel到DataTable
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static DataTable ReadExcel(string excelFileName, int sheetIndex = 0, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetType(excelFileName);
            return new ExcelReader(excelFileName, type).ReadSheet(sheetIndex, columnLength);
        }


        /// <summary>
        /// 读取excel到IEnumerable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static DataTable ReadExcel(Stream stream, string contentType, int sheetIndex = 0, int columnLength = 11)
        {
            var type = NPOITypeUtils.GetTypeByContentType(contentType);
            return new ExcelReader(stream, type).ReadSheet(sheetIndex, columnLength);
        }

    }
}
