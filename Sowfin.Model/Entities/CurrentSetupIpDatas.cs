using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class CurrentSetupIpDatas
    {
        public long Id { get; set; }
        public string? LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string? Unit { get; set; }
        public long? CurrentSetupId { get; set; }
        public List<CurrentSetupIpValues>? CurrentSetupIpValues { get; set; }
    }

    public class CurrentSetupIpValues
    {
        public long Id { get; set; }
        public long? CurrentSetupIpDatasId { get; set; }
        public string? Year { get; set; }
        public string? Value { get; set; }
    }
}
