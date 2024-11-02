using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{
    public class CurrentSetupIpFillingsViewModel
    {
        public string StatementType { get; set; }
        public List<CurrentSetupIpDatasViewModel> CurrentSetupIpDatasViewModelVM { get; set; }
    }
    public class CurrentSetupIpDatasViewModel
    {
        public long Id { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string Unit { get; set; }
        public long? CurrentSetupId { get; set; }
        public List<CurrentSetupIpValuesViewModel> CurrentSetupIpValuesVM { get; set; }
    }
    public class CurrentSetupIpValuesViewModel
    {
        public long Id { get; set; }
        public long? CurrentSetupIpDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
