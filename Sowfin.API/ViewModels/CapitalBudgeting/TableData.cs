using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalBudgeting
{
    public class TableData
    {
        public object[][][] RevenueVariableCostTier { get; set; }
        public object[][][] CapexDepreciation { get; set; }
        public object[][][] OtherFixedCost { get; set; }
        public object[][] WorkingCapital { get; set; }
    }
    public class TableData1
    {
        public object[][][] RevenueVariableCostTier { get; set; }
        public object[][][] CapexDepreciation { get; set; }
        public object[][][] OtherFixedCost { get; set; }
        public object[][] WorkingCapital { get; set; }
    }
}
