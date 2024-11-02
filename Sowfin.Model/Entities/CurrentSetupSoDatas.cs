using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class CurrentSetupSoDatas
    {
        public long Id { get; set; }
        public string? LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string? Unit { get; set; }
        public long? CurrentSetupId { get; set; }
        public List<CurrentSetupSoValues>? CurrentSetupSoValues { get; set; }
    }

    public class CurrentSetupSoValues
    {
        public long Id { get; set; }
        public long? CurrentSetupSoDatasId { get; set; }
        public string? Year { get; set; }
        public string? Value { get; set; }
    }
}
