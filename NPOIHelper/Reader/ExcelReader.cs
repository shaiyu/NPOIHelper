using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace NPOIHelper
{
    public class ExcelReader
    {
        /// <summary>
        /// 当前工作簿
        /// </summary>
        public IWorkbook Workbook { get; }
        /// <summary>
        /// Excel类型 xls/xlsx
        /// </summary>
        public NPOIType Type { get; }
        /// <summary>
        /// Excel文件地址
        /// </summary>
        //public string FileName { get; }

        internal ExcelReader(string fileName, NPOIType type)
        {
            Type = type;
            Workbook = ReadFile(fileName);
        }

        internal ExcelReader(Stream stream, NPOIType type)
        {
            Type = type;
            Workbook = ReadFile(stream);
        }

        private IWorkbook ReadFile(string fileName)
        {
            //初始化信息
            try
            {
                //using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                //开启共享锁，即使被占用也能打开
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    return ReadFile(file);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private IWorkbook ReadFile(Stream stram)
        {
            IWorkbook workbook;
            //初始化信息
            if (Type == NPOIType.xlsx)
            {
                //workbook = new XSSFWorkbook(fileName);
                workbook = new XSSFWorkbook(stram);
            }
            else
            {
                workbook = new HSSFWorkbook(stram);
            }
            return workbook;
        }

        public IEnumerable<T> ReadSheet<T>(int sheetIndex, int maxColumnLength = 11)
        {
            var reader = new ListSheetReader(GetSheet(sheetIndex), Type, maxColumnLength);
            return reader.ReadSheet<T>();
        }

        public DataTable ReadSheet(int sheetIndex, int maxColumnLength = 11)
        {
            var reader = new DataTableSheetReader(GetSheet(sheetIndex), Type, maxColumnLength);
            return reader.ReadSheet();
        }

        //public List<ISheet> Sheets => GetSheets().ToList();

        //public IEnumerable<ISheet> GetSheets()
        //{
        //    for (int k = 0; k < Workbook.NumberOfSheets; k++)
        //    {
        //        ISheet sheet = Workbook.GetSheetAt(k);
        //        if (sheet.PhysicalNumberOfRows <= 0) //实际行数，判断是否为空工作表
        //        {
        //            continue;
        //        }
        //        yield return sheet;
        //    }
        //}


        public ISheet GetSheet(int sheetIndex)
        {
            return Workbook.GetSheetAt(sheetIndex);
        }
    }
}
