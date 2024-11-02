using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
      public  class InitialSetup_IValuation
    {
        public long Id { get; set; }
        public string CIKNumber { get; set; }
        public string Company { get; set; }
        public int? YearFrom { get; set; }
        public int? YearTo { get; set; }
        public long? SourceId { get; set; }
        public long UserId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public int? ExplicitYearCount { get; set; }
        public bool IsActive { get; set; }

    }
}
