using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
     public class TaxRates_IValuation
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
