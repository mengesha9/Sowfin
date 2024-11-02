using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class CategoryByInitialSetupViewModel
    {
        public long Id { get; set; }
        public long? DatasId { get; set; }
        public long? InitialSetupId { get; set; }
        public string Category { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
