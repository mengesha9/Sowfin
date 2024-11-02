using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{
    public class PayoutPolicy_ScenarioOutputFillingsViewModel
    {
        public string StatementType { get; set; }
        public List<PayoutPolicy_ScenarioOutputDatasViewModel> PayoutPolicy_ScenarioOutputDatasVM { get; set; }
    }
    public class PayoutPolicy_ScenarioOutputDatasViewModel
    {
        public long Id { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string Unit { get; set; }
        public long? CurrentSetupId { get; set; }
        public List<PayoutPolicy_ScenarioOutputValuesViewModel> PayoutPolicy_ScenarioOutputValuesVM { get; set; }
    }

    public class PayoutPolicy_ScenarioOutputValuesViewModel
    {
        public long Id { get; set; }
        public long? PayoutPolicy_ScenarioOutputDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
