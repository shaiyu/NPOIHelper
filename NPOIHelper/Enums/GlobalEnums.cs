using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Text;

namespace NPOIHelper.Enums
{
    public static class GlobalEnums
    {
        /// <summary>
        /// 获取枚举描述性数据
        /// </summary>
        /// <param name="enum"></param>
        /// <returns></returns>
        public static string GetDescription(this Enum @enum)
        {
            Type type = @enum.GetType();
            FieldInfo field = type.GetField(@enum.ToString());
            if (field == null)
                return string.Empty;
            var attrs = field.GetCustomAttribute<DescriptionAttribute>();
            if (attrs == null)
                return string.Empty;
            else
                return attrs.Description;
        }

        /// <summary>
        /// 获取枚举携带的值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enum"></param>
        /// <returns></returns>
        public static T GetDefaultValue<T>(this Enum @enum)
        {
            Type type = @enum.GetType();
            FieldInfo field = type.GetField(@enum.ToString());
            if (field == null)
                return default(T);
            var attrs = field.GetCustomAttribute<DefaultValueAttribute>();
            if (attrs == null)
                return default(T);
            else
                return attrs.Value == null ? default(T) : (T)attrs.Value;
        }
    }

    public enum NPOIType
    {
        [DefaultValue("application/vnd.ms-excel")]
        xls,
        [DefaultValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
        xlsx
    }

}
