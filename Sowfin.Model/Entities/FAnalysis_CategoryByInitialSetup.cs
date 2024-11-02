using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{

    public class FAnalysis_CategoryByInitialSetup
    {
        public long Id { get; set; }
        public long? DatasId { get; set; }
        public long? FAnalysis_InitialSetupId { get; set; }
        public string Category { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }

}
