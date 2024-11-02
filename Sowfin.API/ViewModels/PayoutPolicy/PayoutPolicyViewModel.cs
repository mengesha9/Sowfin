using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{

    public class PayoutPolicyViewModel
    {
        public long Id { get; set; }
        public double currentSharePrice { get; set; }
        public double NumberSharesBasic { get; set; }
        public double NumberSharesDiluted { get; set; }
        public double CostOfEquity { get; set; }
        public double PreferredSharePrice { get; set; }
        public double NoPreferredShares { get; set; }
        public double PreferredDividend { get; set; }
        public double CostOfPrefEquity { get; set; }
        public double MarketValueDebt { get; set; }
        public double CostOfDebt { get; set; }
        public double CashNeededWorkingCap { get; set; }
        public double InterestCoverRatio { get; set; }
        public double InterestIncomeRate { get; set; }
        public double MarginalTax { get; set; }
        public double FreeCashFlow { get; set; }
        public string SummaryOutput { get; set; }
        public double TotalOngoingPayout { get; set; }
        public double OneTimePayout { get; set; }
        public double StockBuyBack { get; set; }
        public double NetSales { get; set; }
        public double Ebitda { get; set; }
        public double DepreciationAmortization { get; set; }
        public double InterestIncome { get; set; }
        public double Ebit { get; set; }
        public double InterestExpense { get; set; }
        public double Ebt { get; set; }
        public double Taxes { get; set; }
        public double NetEarings { get; set; }
        public double CashEquivalent { get; set; }
        public double TotalCurrentAssets { get; set; }
        public double TotalAssests { get; set; }
        public double TotalDebt { get; set; }
        public double ShareholderEquity { get; set; }
        public double CashFlowOperation { get; set; }
        public int SummaryFlag { get; set; }
        public int ApprovalFlag { get; set; }
        public string ScenarioObject { get; set; }
        public string ScenarioSummary { get; set; }
        public int ScenarioFlag { get; set; }
        public long UserId { get; set; }
        public string MarketValueDebtUnit { get; set; } ///debt unit
        public string CashCapitalUnit { get; set; }  /// Cash Needed for Working Capital  other inputs
        public string CashFlowUnit { get; set; }  ///Cash Flow from Operations ($M)
        public EquityUnits equityUnits
        {
            get; set;
        }

        [NotMapped]
        public PreferredEquityUnits prefferedEquityUnits
        {
            get; set;

        }
        [NotMapped]
        public PayoutPolicyUnits payoutPolicyUnits
        {
            get; set;
        }
        [NotMapped]
        public IncomeStatementUnits incomeStatementUnits
        {
            get; set;
        }
        [NotMapped]
        public BalanceSheetUnits balanceSheetUnits
        {
            get; set;
        }


    }
}