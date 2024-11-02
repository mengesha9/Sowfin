using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CostOfCapital
{
    public class Methods
    {
        public class Method1
        {
            public double RiskFreeRate { get; set; }
            public double DefaultSpread { get; set; }
        }

        public class Method2
        {
            public double YeildToMaturity { get; set; }
            public double ProbabilityOfDefault { get; set; }
            public double ExpectedLossRate { get; set; }
        }
    }
}
