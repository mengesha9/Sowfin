using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class Interest_IValuationViewModel
    {
        public long Id { get; set; }
        public string Interest_Income { get; set; }
        public string Interest_Expense { get; set; }
        public long? InitialSetupId { get; set; }
        public long UserId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public string TableData { get; set; }
        public int Yearcount { get; set; }

    }
}
