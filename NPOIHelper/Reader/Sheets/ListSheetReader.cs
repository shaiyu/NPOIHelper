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
    public class ListSheetReader : SheetReader
    {
        readonly Type dateTimeType = typeof(DateTime);
        readonly Type nullableDateTimeType = typeof(DateTime?);
        readonly Type guidType = typeof(Guid);
        readonly Type nullableGuidType = typeof(Guid?);
        readonly Type stringType = typeof(string);

        public ListSheetReader(ISheet sheet, NPOIType type, int columnLength = 11) : base(sheet, type, columnLength)
        {
        }

        public IEnumerable<T> ReadSheet<T>()
        {
            if (!HasData())
            {
                yield break;
            }

            //对应Excel表格中的属性信息
            PropertyInfo[] properties = null;
            IRow row;
            bool isFirstRow = true;

            IEnumerator rows = Sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;

                if (isFirstRow)
                {
                    isFirstRow = false;
                    //每个表格第一行，标题行
                    //将第一列作为列表头
                    if (properties == null || properties.Length == 0)
                    {
                        //并确定列数
                        //maxcols = sheet.GetRow(0).LastCellNum;
                        //row.LastCellNum --> colCount --> maxcols
                        ColumnLength = row.Cells.Count > ColumnLength ? ColumnLength : row.Cells.Count;
                        properties = GetSortProperties<T>(row);
                    }
                }
                else
                {
                    T t = ReadRow<T>(properties, row); ;
                    yield return t;
                }
            }
        }

        /// <summary>
        /// 读行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="properties"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private T ReadRow<T>(PropertyInfo[] properties, IRow row)
        {
            T t = Activator.CreateInstance<T>();
            //读列
            for (int i = 0; i < ColumnLength; i++)
            {
                //表格
                ICell cell = row.GetCell(i);
                //i列对应的属性信息
                PropertyInfo prop = properties[i];

                object value = ReadColumn(cell);

                try
                {
                    //设置t的属性值为value
                    SetValue(t, prop, value);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex);
                }
            }
            return t;
        }

        /// <summary>
        /// 设置属性值为value
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <param name="prop"></param>
        /// <param name="value"></param>
        private void SetValue<T>(T t, PropertyInfo prop, object value)
        {
            if (value == null || prop == null || !prop.CanWrite)
            {
                return;
            }

            var valueType = value.GetType();
            //DateTime类型
            if (prop.PropertyType == nullableDateTimeType || prop.PropertyType == dateTimeType)
            {
                //value是DateTime类型, 直接赋值
                if (valueType == dateTimeType)
                {
                    prop.SetValue(t, value);
                }
                //否则, 尝试转换成DateTime
                else if (valueType == stringType && DateTime.TryParse(value.ToString(), out DateTime dateTimeValue))
                {
                    prop.SetValue(t, dateTimeValue);
                }
            }
            //Guid类型
            else if ((prop.PropertyType == nullableGuidType || prop.PropertyType == guidType) && valueType == stringType)
            {
                //尝试转换成Guid
                if (Guid.TryParse(value.ToString(), out Guid guidValue))
                {
                    prop.SetValue(t, guidValue);
                }
            }
            else
            {
                object _value = Convert.ChangeType(value, prop.PropertyType);
                prop.SetValue(t, _value);
            }
        }

        /// <summary>
        /// 获取和Row.Column位置一一对应的PropertyInfo
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        private PropertyInfo[] GetSortProperties<T>(IRow row)
        {
            Type type = typeof(T);
            PropertyInfo[] properties = type.GetProperties();
            PropertyInfo[] sortProperties = new PropertyInfo[ColumnLength];

            for (int i = 0; i < ColumnLength; i++)
            {
                var title = row.GetCell(i) + "";
                PropertyInfo titleProp = null;

                if (!string.IsNullOrWhiteSpace(title))
                {
                    foreach (var prop in properties)
                    {
                        var attr = prop.GetCustomAttribute<TitleNameAttribute>();
                        if (attr != null && title == attr.Value + "")
                        {
                            titleProp = prop;
                            break;
                        }
                    }
                }
                sortProperties[i] = titleProp;
            }
            return sortProperties;
        }
    }
}
