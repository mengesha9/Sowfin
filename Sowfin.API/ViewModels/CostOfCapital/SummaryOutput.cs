using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CostOfCapital
{
    public class SummaryOutput
    {
        public double AdjustedeBeta { get; set; }
        public double MarketRiskPremium { get; set; }
        public double Costofequity { get; set; }
        public double CostofpreferredEquity { get; set; }
        public double CostOfDebtMethod { get; set; }
        public double Unleveredcostofcaptial { get; set; }
        public double WeightedAverage { get; set; }
        public double AdjustedWACC { get; set; }
    }
}
