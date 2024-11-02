using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalBudgeting
{
    public class SensiBody
    {
        public long Id { get; set; }
        public long UserId { get; set; }
        public long ProjectId { get; set; }
        public string SensitivitySnapShot { get; set; }
        public string Description { get; set; }
        public string ChangeNpv { get; set; }
        public string Npv { get; set; }
        public string SnapShot { get; set; }

    }
}
