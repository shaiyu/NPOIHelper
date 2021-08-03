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
    internal class DataTableExcelReader : ExcelReader
    {
        public DataTableExcelReader(string fileName, NPOIType type, int columnLength = 11) : base(fileName, type, columnLength)
        {
        }
        public DataTableExcelReader(Stream stream, NPOIType type, int columnLength = 11) : base(stream, type, columnLength)
        {
        }
        public DataTable Read(int? sheetIndex = null)
        {
            DataTable dt = new();
            IRow row;
            bool isFirstRow;
            bool isAddTitle = false;

            for (int i = 0; i < Sheets.Count; i++)
            {
                if (sheetIndex != null && sheetIndex != i)
                {
                    continue;
                }
                var sheet = Sheets[i];

                isFirstRow = true;
                IEnumerator rows = sheet.GetRowEnumerator();
                while (rows.MoveNext())
                {
                    row = (IRow)rows.Current;
                    if (isFirstRow)
                    {
                        isFirstRow = false;
                        //每个表格第一行，标题行
                        //将第一列作为列表头

                        if (!isAddTitle)
                        {
                            //确定列数
                            ColumnLength = row.Cells.Count > ColumnLength ? ColumnLength : row.Cells.Count;
                            for (int j = 0; j < ColumnLength; j++)
                            {
                                //dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());  
                                //将第一列作为列表头
                                dt.Columns.Add(row.GetCell(j) + "");
                            }
                            isAddTitle = true;
                        }
                    }
                    else
                    {
                        DataRow dr = dt.NewRow();
                        ReadRow(ref dr, row, ColumnLength);
                        dt.Rows.Add(dr);
                    }
                }

            }

            return dt;
        }

        public void ReadRow(ref DataRow dr, IRow row, int maxcols = 11)
        {
            //var colCount = row.Cells.Count > maxcols ? maxcols : row.Cells.Count;
            for (int i = 0; i < maxcols; i++)//row.LastCellNum --> colCount --> maxcols
            {
                ICell cell = row.GetCell(i);
                dr[i] = ReadColumn(cell);
            }
        }

    }
}
