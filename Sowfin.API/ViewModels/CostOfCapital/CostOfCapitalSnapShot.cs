using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CostOfCapital
{
    public class CostOfCapitalSnapShot
    {
        public long Id { get; set; }
        public long UserId { get; set; }
        public string SnapShot { get; set; }
        public string Description { get; set; }

    }
}
