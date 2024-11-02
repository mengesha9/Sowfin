using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum ValuationTechniqueEnum
    {
        [Description("DCF (Discounted Cash Flow) using WACC (Weighted Average Cost of Capital) Leverage Policy: Constant Debt - Equity ratio - Same as Firm's Target Leverage Ratio")]
        V1 = 1,
        [Description("DCF (Discounted Cash Flow) using APV (Adjusted Present Value) Leverage Policy: Constant Debt - Equity ratio - Same as Firm's Target Leverage Ratio")]
        V2=2,
        [Description("FTE (Flow-to-Equity) using Cost of Equity Method Leverage Policy: Constant Debt - Equity ratio - Same as Firm's Target Leverage Ratio")]
        V3=3,
        [Description("DCF (Discounted Cash Flow) using Project Based Cost of Capital Leverage Policy: Constant Debt-Equity ratio (Target Leverage Ratio) - Different from Rest of the Firm")]
        V4=4,
        [Description("")]
        v5=5,
        [Description("")]
        v6 = 6,
        [Description("")]
        v7 = 7,
        [Description("")]
        v8 = 8
    }
}
