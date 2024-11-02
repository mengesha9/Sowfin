using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum StatementTypeEnum
    {
        [Description("INCOME")]
        IncomeStatement = 1,
        [Description("BALANCE SHEET")]
        BalanceSheet = 2,
        [Description("CASH FLOW STATEMENT")]
        CashFlowStatement = 3,
        [Description("RETAINED EARNINGS STATEMENT")]
        RetainedEarningsStatement = 4,
        [Description("PAYOUT POLICY FORECAST")]
        PayoutPolicyForcast = 5,
        [Description("NOPLAT CALCULATIONS")]
        NoplatCalculations = 6,
        [Description("INVESTED CAPITAL CALCULATIONS")]
        InvestedCapitalCalculations = 7,
        [Description("Non-Operating Assets")]
        NonOperatingAssets = 8,
        [Description("NonEquity Claims")]
        NonEquityClaims = 9,
        [Description("ROIC")]
        ROIC = 10,
        [Description("Free Cash Flow")]
        FCF = 11,
        [Description("Discounted Cash Flow")]
        DCF1 = 12,
        [Description("Discounted Cash Flow")]
        DCF2 = 13,
        [Description("Discounted Cash Flow")]
        DCF3 = 14,
        [Description("Valuation Summary")]
        Valuation = 15,
        [Description("Sources of Financing")]
        SourcesOfFinancing = 16,
        [Description("Other Inputs")]
        OtherInputs = 17,
        [Description("Current Payout Policy")]
        CurrentPayoutPolicy = 18,
        [Description("Financial Statements: Pre-Payout")]
        FinancialStatementsPrePayout = 19,

        [Description("Current Capital Structure")]
        CurrentCapitalStructure = 20,
        [Description("Current Cost of Capital")]
        CurrentCostOfCapital = 21,
        [Description("Leverage Ratios")]
        LeverageRatios = 22,
        [Description("Cash Balance: Post-Payout")]
        CashBalancePostPayout = 23,
        [Description("Internal Valuation: Post-Payout")]
        InternalValuationPostPayout = 24,
        [Description("Current Payout Analysis")]
        CurrentPayoutAnalysis = 25,
        [Description("Shares Outstanding & EPS: Post-Payout")]
        SharesOutstandingAndEPSPostPayout = 26,
        [Description("Financial Statements: Post-Payout")]
        FinancialStatementsPostPayout = 27,
        [Description("Debt Ratios & Analysis: Post-Payout")]
        DebtRatiosAndAnalysisPostPayout = 28,
        [Description("Inputs")]
        Inputs=29,
        [Description("New Payout Policy")]
        NewPayoutPolicy=30
    }
}
