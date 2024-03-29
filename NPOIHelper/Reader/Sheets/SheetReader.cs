﻿using NPOI.HSSF.UserModel;
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
    public abstract class SheetReader
    {
        /// <summary>
        /// 当前工作簿
        /// </summary>
        public IWorkbook Workbook { get; }
        public ISheet Sheet { get; }
        /// <summary>
        /// Excel类型 xls/xlsx
        /// </summary>
        public NPOIType Type { get; }
        /// <summary>
        /// Excel列数
        /// </summary>
        public int ColumnLength { get; protected set; } = 11;

        public SheetReader(ISheet sheet, NPOIType type, int columnLength = 11)
        {
            Workbook = sheet.Workbook;
            Sheet = sheet;
            Type = type;
            ColumnLength = columnLength;
        }

        public bool HasData()
        {
            return Sheet != null && Sheet.PhysicalNumberOfRows > 0;
        }

        public object ReadColumn(ICell cell) {

            if (cell == null)
            {
                return null;
            }

            try
            {
                return ReadColumnByCellType(cell);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            return null;
        }

        private object ReadColumnByCellType(ICell cell)
        {
            object value = null;

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
