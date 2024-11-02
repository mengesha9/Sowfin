using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{
    public class PayoutScenario
    {
        public long Id { get; set; }
        public double TargetDebtToEquity { get; set; }
        public double CostOfDebt { get; set; }
        public double TotalPayoutAnnual { get; set; }
        public double OneTimePayout { get; set; }
        public double StockBuyBack { get; set; }
        public string TotalPayoutAnnualUnit { get; set; }
        public string OneTimePayoutUnit { get; set; }
        public string StockBuyBackUnit { get; set; }
        public long UserId { get; set; }
    }
}
