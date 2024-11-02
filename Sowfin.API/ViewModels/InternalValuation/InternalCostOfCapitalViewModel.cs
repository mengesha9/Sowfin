using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class InternalCostOfCapitalViewModel
    {
        public long Id { get; set; }
        public decimal? RiskFreeRate { get; set; }
        public decimal? CostOfDebt { get; set; }
        public decimal? WeightedAverage { get; set; }
        public long? CompanyId { get; set; }
        public long UserId { get; set; }
        public long InitialSetupId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
