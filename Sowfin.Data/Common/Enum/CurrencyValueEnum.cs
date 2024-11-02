using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum CurrencyValueEnum
    {

        [Description("$")]
        Dollar = 1,
        [Description("$K")]
        DollarK = 2,
        [Description("$M")]
        DollarM = 3,
        [Description("$B")]
        DollarB = 4,
        [Description("$T")]
        DollarT = 5
    }
}
