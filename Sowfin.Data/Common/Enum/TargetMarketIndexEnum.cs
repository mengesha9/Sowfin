using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum TargetMarketIndexEnum
    {
        [Description("S&P 500")]
        SNP500 = 1,
        [Description(" S&P 1500")]
        SNP1500 = 2,
        [Description("Russell 1000")]
        Russell1000 = 3,
        [Description("Russell 3000")]
        Russell3000 = 4,
        [Description("MSCI USA IMI")]
        MSCIUSAIMI = 5
    }
}
