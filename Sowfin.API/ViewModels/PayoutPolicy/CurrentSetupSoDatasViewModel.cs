using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{

    public class CurrentSetupSoFillingsViewModel
    {
        public string StatementType { get; set; }
        public List<CurrentSetupSoDatasViewModel> CurrentSetupSoDatasViewModelVM { get; set; }
    }
    public class CurrentSetupSoDatasViewModel
    {
        public long Id { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string Unit { get; set; }
        public long? CurrentSetupId { get; set; }
        public List<CurrentSetupSoValuesViewModel> CurrentSetupSoValuesVM { get; set; }
    }
    public class CurrentSetupSoValuesViewModel
    {
        public long Id { get; set; }
        public long? CurrentSetupSoDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
