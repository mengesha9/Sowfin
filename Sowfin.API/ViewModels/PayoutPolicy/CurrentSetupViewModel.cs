using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{
    public class CurrentSetupViewModel
    {
        public long Id { get; set; }
        public string CurrentYear { get; set; }
        public string CIKNumber { get; set; }
        public string Company { get; set; }
        public string EndYear { get; set; }
        public long? InitialSetupId { get; set; }
        public long? UserId { get; set; }
        public bool? isChanged { get; set; }
    }
}
