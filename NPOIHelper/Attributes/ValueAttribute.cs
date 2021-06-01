using System;
using System.Collections.Generic;
using System.Text;

namespace NPOIHelper
{

    public class ValueAttribute : Attribute
    {
        public object Value { get; set; }
        public ValueAttribute(string value)
        {
            Value = value;
        }
    }


}
