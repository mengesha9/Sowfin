using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static Sowfin.Model.Entities.CapitalStructure;

namespace Sowfin.API.ViewModels
{
    public class CapitalStructureViewModel
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
        public int SummaryFlag { get; set; }
        public string CapitalPieChart { get; set; }
        public string ScenarioPieChart { get; set; }
        public Equity equity { get; set; }
        public EquityUnit equityUnits { get; set; }
        public PrefferedEquity prefferedEquity { get; set; }
        public PrefferedEquityUnit prefferedEquityUnit { get; set; }
        public Debt debt { get; set; }
        public DebtUnit debtUnit { get; set; }


    }
}
