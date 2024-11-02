using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum DataSourceEnum
    {
        [Description("Excel Upload")]
        ExcelUpload = 1,
        [Description("Bloomberg")]
        Bloomberg = 2,
        [Description("Thomson Reuters")]
        ThomsonReuters = 3,
        [Description("Morning Star")]
        MorningStar = 4
    }
}
