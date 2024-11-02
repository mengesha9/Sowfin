using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum HeadersEnum
    {
        [Description("Source of Financing")]
        SourceOfFinancing = 1,
        [Description("Other Inputs")]
        OtherInputs = 2,
        [Description("Current Capital Structure")]
        CurrentCapitalStructure = 3,
        [Description("Current Cost of Capital")]
        CurrentCostofCapital = 4,
        [Description("Cash Balance")]
        CashBalance = 5,
        [Description("Leverage Ratios")]
        LeverageRatios = 6,
        [Description("Internal Valuation")]
        InternalValuation = 7,
        [Description("Cost of Equity Capital")]
        CostofEquityCapital = 8,
        [Description("Cost of Preferred Equity")]
        CostofPreferredEquity = 9,
        [Description("Cost of Debt Capital")]
        CostofDebtCapital = 10,

    }

    public enum BetaSourceEnum
    {
        [Description("Calculate Beta")]
        CalculateBeta = 1,
        [Description("External Source")]
        ExternalSource = 2,
        [Description("Manual Entry")]
        ManualEntry=3
    }

    public enum ScenarioAnalysisHeadersEnum
    {
        [Description("Source of Financing")]
        SourceOfFinancing = 1,
        [Description("Other Inputs")]
        OtherInputs = 2,
        [Description("New Capital Structure")]
        NewCapitalStructure = 3,
        [Description("New Cost of Capital")]
        NewCostofCapital = 4,
        [Description("New Cash Balance")]
        NewCashBalance = 5,
        [Description("New Leverage Ratios")]
        NewLeverageRatios = 6,
        [Description("New Internal Valuation")]
        NewInternalValuation = 7,
        [Description("Cost of Equity Capital")]
        CostofEquityCapital = 8,
        [Description("Cost of Preferred Equity")]
        CostofPreferredEquity = 9,
        [Description("Cost of Debt Capital")]
        CostofDebtCapital = 10,

    }
}