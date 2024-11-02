using System;
using System.Collections.Generic;
using System.Text;


using System;
using System.Collections.Generic;

namespace Sowfin.Model.Entities
{
    public class PayoutPolicy_ScenarioOutputDatas
    {
        public long Id { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string Unit { get; set; }
        public long? CurrentSetupId { get; set; }
        public List<PayoutPolicy_ScenarioOutputValues> PayoutPolicy_ScenarioOutputValues { get; set; } 
    }

    public class PayoutPolicy_ScenarioOutputValues
    {
        public long Id { get; set; }
        public long? PayoutPolicy_ScenarioOutputDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
