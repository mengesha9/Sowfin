using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class Project_SnapshotDatas
    {
        public long Id { get; set; }
        public long? ProjectId { get; set; }
        public long? HeaderId { get; set; }
        public string SubHeader { get; set; }
        public string LineItem { get; set; }
        public int? UnitId { get; set; }
        public int? ValueTypeId { get; set; }
        public bool HasMultiYear { get; set; }
        public double? Value { get; set; }
    }

    public class Project_SnapshotValues
    {
        public long Id { get; set; }
        public long? Project_SnapshotDatasId { get; set; }
        public int? Year { get; set; }
        public double? Value { get; set; }
    }
}
