using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalStructure
{
    public class ScenarioSummaryOutput
    {
        public string equityValue { get; set; }
        public string preferredEquityValue { get; set; }
        public string debtValue { get; set; }
        public string changeInDebt { get; set; }
        public string netDebtValue { get; set; }
        public string cashEquivalent { get; set; }
        public string excessCash { get; set; }
        public double costOfEquity { get; set; }
        public double costOfPreferredEquity { get; set; }
        public double weightedAverage { get; set; }
        public double costOfDebt { get; set; }
        public double unleveredCostOfCapital { get; set; }
        public double debtToEquity { get; set; }
        public double debtToValue { get; set; }
    }
}
