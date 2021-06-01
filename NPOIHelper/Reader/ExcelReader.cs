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
    internal abstract class ExcelReader
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
        public string FileName { get; }
        /// <summary>
        /// Excel列数
        /// </summary>
        public int ColumnLength { get; protected set; } = 11;

        public ExcelReader(string fileName, int columnLength = 11)
        {
            FileName = fileName;

            Type = GetType(fileName);
            Workbook = ReadFile(fileName);
            ColumnLength = columnLength;
        }

        public IWorkbook ReadFile(string fileName)
        {
            IWorkbook workbook;
            //初始化信息
            try
            {
                //using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                //开启共享锁，即使被占用也能打开
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (Type == NPOIType.xlsx)
                    {
                        //workbook = new XSSFWorkbook(fileName);
                        workbook = new XSSFWorkbook(file);
                    }
                    else
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return workbook;
        }

        public NPOIType GetType(string fileName)
        {
            var ext = Path.GetExtension(fileName);
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


        public IEnumerable<ISheet> Sheets
        {
            get
            {
                for (int k = 0; k < Workbook.NumberOfSheets; k++)
                {
                    ISheet sheet = Workbook.GetSheetAt(k);
                    if (sheet.PhysicalNumberOfRows <= 0) //实际行数，判断是否为空工作表
                    {
                        continue;
                    }
                    yield return sheet;
                }
            }
        }

        public object ReadColumn(ICell cell)
        {
            object value = null;
            if (cell == null)
            {
                return value;
            }

            //读取Excel格式，根据格式读取数据类型
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型处理
                    break;
                case CellType.String: //字符串类型
                    value = cell.StringCellValue;
                    break;
                case CellType.Numeric: //数字类型 - >电话号码      
                    //NPOI中数字和日期都是NUMERIC类型的，这里对其进行判断是否是日期类型
                    if (DateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        value = cell.DateCellValue;
                    }
                    else//其他数字类型
                    {
                        value = cell.NumericCellValue;
                    }

                    //cell.SetCellType(CellType.String);
                    //value = cell.StringCellValue;
                    break;
                case CellType.Formula:
                    IFormulaEvaluator e = Type == NPOIType.xlsx ?
                        new XSSFFormulaEvaluator(Workbook) : new HSSFFormulaEvaluator(Workbook);
                    value = e.Evaluate(cell).StringValue;
                    break;
                default:
                    value = cell.StringCellValue;
                    break;
            }
            return value;
        }
    }
}
