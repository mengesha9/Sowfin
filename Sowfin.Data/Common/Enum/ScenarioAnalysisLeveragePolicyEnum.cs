using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum ScenarioAnalysisLeveragePolicyEnum
    {
        [Description("Constant Debt-to-Equity Ratio (Target Leverage Ratio)")]
        SADebtToEquityRatio = 1,
        [Description("Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")]
        SAAnnuallyAdjustDebt = 2,
        [Description("Constant Permanent Debt")]
        SAConstantPermanentDebt = 3,
        [Description("Constant Interest Coverage Ratio (% of Free Cash Flow)")]
        SAConstantInterestRatio = 4
    }
}
