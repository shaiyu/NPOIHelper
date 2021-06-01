using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NPOIHelper
{
    public enum NPOIType
    {
        [DefaultValue("application/vnd.ms-excel")]
        xls,
        [DefaultValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
        xlsx
    }
}
