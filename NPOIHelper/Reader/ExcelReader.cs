using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
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
    public class ExcelReader : IDisposable
    {
        private bool disposedValue;

        /// <summary>
        /// 当前工作簿
        /// </summary>
        public IWorkbook Workbook { get; private set; }
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

        private IWorkbook ReadFile(Stream stream)
        {
            IWorkbook workbook;
            //初始化信息
            if (Type == NPOIType.xlsx)
            {
                //workbook = new XSSFWorkbook(fileName);
                workbook = new XSSFWorkbook(stream);
            }
            else
            {
                workbook = new HSSFWorkbook(stream);
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

        public void Close()
        {
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)+

                    if (Workbook != null)
                    {
                        Workbook.Close();
                    }
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                Workbook = null;
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ExcelReader()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
