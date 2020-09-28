using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelper
{
    public class TypeFactory
    {

        /// <summary>
        /// 读入配置映射
        /// key enumType 
        /// value ClassType
        /// </summary>
        public void LoadConfig()
        {

        }

        /// <summary>
        /// 简单配置
        /// </summary>
        /// <param name="columnType"></param>
        /// <returns></returns>
        public static IDataType Get(ColumnType columnType)
        {
            switch (columnType)
            {
                case ColumnType.Number:
                    return new NumberDataType(ColumnType.Number);
                case ColumnType.NumDecimal0:
                    return new Number0DataType(ColumnType.NumDecimal0);
                case ColumnType.NumDecimal2:
                    return new Number2DataType(ColumnType.NumDecimal2);
                case ColumnType.NumberPercentage:
                    return new NumberPrecentageDataType(ColumnType.NumberPercentage);
                case ColumnType.NumberScientificNotation:
                    return new NumberScientificNotationDataType(ColumnType.NumberScientificNotation);
                case ColumnType.Date:
                    return new DateDataType(ColumnType.Date);
                case ColumnType.DateTime:
                    return new DateTimeDataType(ColumnType.DateTime);
                case ColumnType.DateFile:
                    //throw new NotImplementedException("暂未实现");
                case ColumnType.Default:
                case ColumnType.String:
                default:
                    break;
            }
            return new DefaultDataType(ColumnType.Default);
        }

    }
}
