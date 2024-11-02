using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace Sowfin.Model.Entities
{
    public partial class CapitalStructure
    {

        public long Id { get; set; }
        public string NewLeveragePolicy { get; set; }
        public double CashEquivalent { get; set; }
        public string CashEquivalentUnit { get; set; }
        public double CashNeededCapital { get; set; }
        public string CashNeededCapitalUnit { get; set; }
        public double InterestCoverage { get; set; }
        public double MarginalTaxRate { get; set; }
        public double FreeCashFlow { get; set; }
        public string FreeCashFlowUnit { get; set; }
        public int ApprovalFlag { get; set; }
        public long UserId { get; set; }
        public string SummaryOutput { get; set; }
        public string ScenarioObject { get; set; }
        public string ScenarioOutput { get; set; }
        public string ScenarioPolicy { get; set; }
        public string PermanentDebtUnit { get; set; }
        public string ScenarioFreeUnit { get; set; }
        public  string _Equity { get; set; }
        public string _EquityUnits { get; set; }
        public string _PrefferedEquity { get; set; }
        public string _PrefferedEquityUnits { get; set; }
        public string _Debt { get; set; }
        public string _DebtUnit { get; set; }
        public int SummaryFlag { get; set; }
        public string CapitalPieChart { get; set; }
        public string ScenarioPieChart { get; set; }

        [NotMapped]
        public Equity equity
        {
            get { return _Equity == null ? null : JsonConvert.DeserializeObject<Equity>(_Equity); }
            set { _Equity = JsonConvert.SerializeObject(value); }

        }
        [NotMapped]
        public EquityUnit equityUnits
        {
            get { return _EquityUnits == null ? null : JsonConvert.DeserializeObject<EquityUnit>(_EquityUnits); }
            set { _EquityUnits = JsonConvert.SerializeObject(value); }

        }
        [NotMapped]
        public PrefferedEquity prefferedEquity
        {
            get { return _PrefferedEquity == null ? null : JsonConvert.DeserializeObject<PrefferedEquity>(_PrefferedEquity); }
            set { _PrefferedEquity = JsonConvert.SerializeObject(value); }
        }
        [NotMapped]
        public PrefferedEquityUnit prefferedEquityUnit
        {
            get { return _PrefferedEquityUnits == null ? null : JsonConvert.DeserializeObject<PrefferedEquityUnit>(_PrefferedEquityUnits); }
            set { _PrefferedEquityUnits = JsonConvert.SerializeObject(value); }

        }
        [NotMapped]
        public Debt debt
        {
            get { return _Debt == null ? null : JsonConvert.DeserializeObject<Debt>(_Debt); }
            set { _Debt = JsonConvert.SerializeObject(value); }

        }
        [NotMapped]
        public DebtUnit debtUnit
        {
            get { return _DebtUnit == null ? null : JsonConvert.DeserializeObject<DebtUnit>(_DebtUnit); }
            set { _DebtUnit = JsonConvert.SerializeObject(value); }

        }
        public class Equity
        {
            public double CurrentSharePrice { get; set; }
            public double NumberShareBasic { get; set; }
            public double NumberShareOutstanding { get; set; }
            public double CostOfEquity { get; set; }
        }
        public class PrefferedEquity
        {
            public double PrefferedSharePrice { get; set; }
            public double PrefferedShareOutstanding { get; set; }
            public double PrefferedDividend { get; set; }
            public double CostPreffEquity { get; set; }
        }
        public class Debt
        {
            public double MarketValueDebt { get; set; }
            public double CostOfDebt { get; set; }
        }
        public class EquityUnit
        {
            public string CurrentSharePriceUnit { get; set; }
            public string NumberShareBasicUnit { get; set; }
            public string NumberShareOutstandingUnit { get; set; }
        }
        public class PrefferedEquityUnit
        {
            public string PrefferedSharePriceUnit { get; set; }
            public string PrefferedShareOutstandingUnit { get; set; }
            public string PrefferedDividendUnit { get; set; }
        }
        public class DebtUnit
        {
            public string MarketValueDebtUnit { get; set; }
        }
    }
}
