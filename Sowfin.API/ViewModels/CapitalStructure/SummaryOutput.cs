using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalStructure
{
    public class SummaryOutput
    {
        public string marketValueEquity { get; set; }
        public string marketValuePreferredEquity { get; set; }
        public string marketValueDebt { get; set; }
        public string marketValueNetDebt { get; set; }
        public double costEquity { get; set; }
        public double costPreferredEquity { get; set; }
        public double costDebt { get; set; }
        public double unleveredCostCapital { get; set; }
        public double weighedAverageCapital { get; set; }
        public string cashEquivalent { get; set; }
        public string excessCash { get; set; }
        public double debtEquityRatio { get; set; }
        public double debtValueRatio { get; set; }
        public double unleveredEnterpriseValue { get; set; }
        public double leverateEnterpriseValue { get; set; }
        public double equityValue { get; set; }
        public double interestTaxShield { get; set; }
        public string stockPrice { get; set; }
    }
}
