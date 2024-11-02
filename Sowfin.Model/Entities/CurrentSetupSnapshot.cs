using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class CurrentSetupSnapshot
    {
        public long Id { get; set; }
        public string? Description { get; set; }
        public long UserId { get; set; }
        public List<CurrentSetupSnapshotDatas>? CurrentSetupSnapshotDatas { get; set; }
    }

    public class CurrentSetupSnapshotDatas
    {
        public long Id { get; set; }
        public string? LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string? Unit { get; set; }
        public long? CurrentSetupSnapshotId { get; set; }
        public List<CurrentSetupSnapshotValues>? CurrentSetupSnapshotValues { get; set; }
    }

    public class CurrentSetupSnapshotValues
    {
        public long Id { get; set; }
        public long? CurrentSetupSnapshotDatasId { get; set; }
        public string? Year { get; set; }
        public string? Value { get; set; }
    }
}