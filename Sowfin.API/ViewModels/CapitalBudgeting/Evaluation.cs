using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalBudgeting
{
    public class Evaluation
    {
 

        public int id { get; set; }
        public int startingYear { get; set; }
        public int noOfYears { get; set; }
        public double discountRate { get; set; }
        public double marginalTaxRate { get; set; }
        public TableData tableData { get; set; }
        public long userId { get; set; }
        public int approvalFlag { get; set; }
        public long projectId { get; set; }
        public int capexCount { get; set; }
        public int revenueCount { get; set; }
        public int fixedCostCount { get; set; }

    }
}
