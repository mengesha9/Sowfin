using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum IntegratedReferenceEnum
    {
        [Description("HistoricalData")]
        HistoricalData = 1,
        [Description("Interest")]
        Interest = 2,
        [Description("Cost Of Capital")]
        CostOfCapital = 3,
        [Description("HistoricalForcastRatio")]
        HistoricalForcastRatio = 4,
        [Description("IntegratedFinancialStmt")]
        IntegratedFinancialStmt = 5,
        [Description("Empty")]
        Empty = 6,
        [Description("PayoutPolicy")]
        PayoutPolicy = 7,
        [Description("Mixed")]
        Mixed = 8
    }
}
