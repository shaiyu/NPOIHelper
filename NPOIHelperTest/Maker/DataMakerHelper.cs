using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NPOIHelperTest
{
    public class DataMakerHelper
    {
        public static T Make<T>()
        {
            var type = typeof(T);
            var properties = type.GetProperties();

            T t = Activator.CreateInstance<T>();

            Type datetimeType = typeof(DateTime);
            Type stringType = typeof(string);
            Type decimalype = typeof(decimal);
            Type longType = typeof(long);
            Type intType = typeof(int);
            Type guidType = typeof(Guid);
            Type enumType = typeof(Enum);
            foreach (var property in properties)
            {
                if (!property.CanWrite)
                {
                    continue;
                }

                var name = property.Name.ToLower();
                var propertyType = property.PropertyType;

                object value = null;
                if (propertyType == guidType)
                {
                    value = Guid.NewGuid();
                }
                else if (propertyType == intType || propertyType == longType)
                {
                    value = new Random().Next(10, 99999);
                }
                else if (propertyType == decimalype)
                {
                    if (name.Contains("rate"))
                    {
                        value = (decimal)new Random().NextDouble();
                    }
                    else
                    {
                        value = (decimal)new Random().Next(10, 99999);
                    }
                }
                else if (propertyType == datetimeType)
                {
                    var second = new Random().Next(0, 10000);
                    value = DateTime.Now.AddSeconds(-second);
                }
                else if (propertyType.BaseType == enumType)
                {
                    var enumFields = propertyType.GetFields().Where(s => s.FieldType == propertyType).ToArray();
                    var enumPosition = new Random().Next(0, enumFields.Length);
                    var enumField = enumFields[enumPosition];
                    value = enumField.GetValue(null);
                }
                else if (propertyType == stringType)
                {
                    if (name.Contains("name"))
                    {
                        value = NameMaker.GetName();
                    }
                    if (name.Contains("phone"))
                    {
                        value = PhoneMaker.Build();
                    }
                }

                property.SetValue(t, value);
            }
            return t;
        }

        public static IEnumerable<T> Makes<T>(long count)
        {
            for (int i = 0; i < count; i++)
            {
                yield return Make<T>();
            }
        }
    }
}
