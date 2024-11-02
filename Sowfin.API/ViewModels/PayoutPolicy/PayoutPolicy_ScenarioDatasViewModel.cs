using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{
    public class PayoutPolicy_ScenarioFillingsViewModel
    {
        public string StatementType { get; set; }
        public List<PayoutPolicy_ScenarioDatasViewModel> PayoutPolicy_ScenarioDatasVM { get; set; }
    }

    public class PayoutPolicy_ScenarioDatasViewModel
    {
        public long Id { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string Unit { get; set; }
        public long? CurrentSetupId { get; set; }
        public List<PayoutPolicy_ScenarioValuesViewModel> PayoutPolicy_ScenarioValuesVM { get; set; }
    }

    public class PayoutPolicy_ScenarioValuesViewModel
    {
        public long Id { get; set; }
        public long? PayoutPolicy_ScenarioDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
