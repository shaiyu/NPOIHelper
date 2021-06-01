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
            return new ListExcelReader(excelFileName, columnLength).Read<T>();
        }

        /// <summary>
        /// 读取Excel到DataTable
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <param name="columnLength"></param>
        /// <returns></returns>
        public static DataTable ReadExcel(string excelFileName, int columnLength = 11)
        {
            return new DataTableExcelReader(excelFileName, columnLength).Read();
        }
    }
}
