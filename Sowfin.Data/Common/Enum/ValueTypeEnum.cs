using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum ValueTypeEnum
    {
        [Description("% Value")]
        percentage = 1,
        [Description("Number Count")]
        Number = 2,
        [Description("Currency Value")]
        Currency = 3,
        [Description("Other")]
        Other = 4
    }
}
