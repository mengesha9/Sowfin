using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum IVSensitivityEnum
    {
        [Description("WACC")]
        WACC = 1,
        [Description("Terminal Growth Rate")]
        TGR = 2,
        [Description("RONIC")]
        RONIC = 3,
        [Description("Net PPE % of Sales")]
        NetPPE =4,
        [Description("COGS % of Sales")]
        COGS = 5,
        [Description("R&D % of Sales")]
        RandD = 6,
        [Description("SG&A % of Sales")]
        SGandA = 7
    }
}
