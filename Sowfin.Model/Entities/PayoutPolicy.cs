using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class PayoutPolicy
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
        public string _EquityUnits { get; set; }
        public string _PreferredUnits { get; set; }
        public string _PayoutUnits { get; set; }
        public string _IncomeUnits { get; set; }
        public string _BalanceUnits { get; set; }
        [NotMapped]
        public EquityUnits equityUnits
        {
            get { return _EquityUnits == null ? null : JsonConvert.DeserializeObject<EquityUnits>(_EquityUnits); }
            set { _EquityUnits = JsonConvert.SerializeObject(value); }
        }

        [NotMapped]
        public PreferredEquityUnits prefferedEquityUnits
        {
            get { return _PreferredUnits == null ? null : JsonConvert.DeserializeObject<PreferredEquityUnits>(_PreferredUnits); }
            set { _PreferredUnits = JsonConvert.SerializeObject(value); }

        }
        [NotMapped]
        public PayoutPolicyUnits payoutPolicyUnits
        {
            get { return _PayoutUnits == null ? null : JsonConvert.DeserializeObject<PayoutPolicyUnits>(_PayoutUnits); }
            set { _PayoutUnits = JsonConvert.SerializeObject(value); }
        }
        [NotMapped]
        public IncomeStatementUnits incomeStatementUnits
        {
            get { return _IncomeUnits == null ? null : JsonConvert.DeserializeObject<IncomeStatementUnits>(_IncomeUnits); }
            set { _IncomeUnits = JsonConvert.SerializeObject(value); }
        }
        [NotMapped]
        public BalanceSheetUnits balanceSheetUnits
        {
            get { return _BalanceUnits == null ? null : JsonConvert.DeserializeObject<BalanceSheetUnits>(_BalanceUnits); }
            set { _BalanceUnits = JsonConvert.SerializeObject(value); }
        }
    }
    public class EquityUnits
    {
        public string CurrentSharePriceUnit { get; set; }
        public string NumberShareBasicUnit { get; set; }
        public string NumberShareOutstandingUnit { get; set; }
    }
    public class PreferredEquityUnits
    {
        public string PrefferedSharePriceUnit { get; set; }
        public string PrefferedShareOutstandingUnit { get; set; }
        public string PrefferedDividendUnit { get; set; }
    }
    public class PayoutPolicyUnits
    {
        public string TotalPayoutUnit { get; set; }
        public string OneTimeDividendUnit { get; set; }
        public string StockAmountUnit { get; set; }
    }
    public class IncomeStatementUnits
    {
        public string NetSalesUnit { get; set; }
        public string EbitaUnit { get; set; }
        public string DeprecAmortizationUnit { get; set; }
        public string IncomeIncomeUnit { get; set; }
        public string EbitUnit { get; set; }
        public string InterestExpenseUnit { get; set; }
        public string EbtUnit { get; set; }
        public string TaxesUnit { get; set; }
        public string NetEarningsUnit { get; set; }
    }
    public class BalanceSheetUnits
    {
        public string CashEquivalentUnit { get; set; }
        public string TotalCurrentAssetUnit { get; set; }
        public string TotalAssetUnit { get; set; }
        public string TotalDebtUnit { get; set; }
        public string ShareEquitUnit { get; set; }
    }

}
