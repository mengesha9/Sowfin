using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalStructure
{
    public class MathClass
    {
        public class ConstantDebtEquity  // this same for 1st two methods
        {
            public double TargetDebtEquity { get; set; }
            public double CostOfDebt { get; set; }
        }

        public class ConstPermanentDebt
        {
            public double ValueOfPermDebt { get; set; }
            public double CostOfDebt { get; set; }
        }

        public class ConstInterestCovRatio
        {
            public double InterestCovRatio { get; set; }
            public double FreeCashFlow { get; set; }
            public double CostOfDebt { get; set; }

        }
    }
}
