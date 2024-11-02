using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class InternalTaxRatesViewModel
    {
        public long Id { get; set; }
        public decimal? Statutory_Federal { get; set; }
        public decimal? Marginal { get; set; }
        public decimal? Operating { get; set; }
        public long? CompanyId { get; set; }
        public long UserId { get; set; }
        public long InitialSetupId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
