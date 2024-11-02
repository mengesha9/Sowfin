using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum LeveragePolicyEnum
    {
        [Description("Constant Debt-to-Equity Ratio  (Target Leverage Ratio)")]
        DebtToEquityRatio = 1,
        [Description("Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")]
        AnnuallyAdjustDebt = 2,
        [Description("Constant Permanent Debt")]
        ConstantPermanentDebt = 3,
        [Description("Constant Interest Coverage Ratio (% of Free Cash Flow)")]
        ConstantInterestRatio = 4
    }
}
