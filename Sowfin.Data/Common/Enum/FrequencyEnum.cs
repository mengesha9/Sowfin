using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
     public enum FrequencyEnum
    {
        [Description("Monthly")]
        Monthly = 1,
        [Description("Weekly")]
        Weekly = 2,
        [Description("Daily")]
        Daily = 3
       
    }
}
