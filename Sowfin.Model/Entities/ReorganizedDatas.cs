using System;
using System.Collections.Generic;

namespace Sowfin.Model.Entities
{
    public class ReorganizedDatas
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public bool?  IsParentItem { get; set; }
        public long? IntegratedDatasId { get; set; }
        public bool IsExplicit_editable { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<ReorganizedValues> ReorganizedValues { get; set; } = new List<ReorganizedValues>();
        public List<Reorganized_ExplicitValues> Reorganized_ExplicitValues { get; set; } = new List<Reorganized_ExplicitValues>();
    }

    public class ReorganizedValues
    {
        public long Id { get; set; }
        public long? ReorganizedDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }

    public class Reorganized_ExplicitValues
    {
        public long Id { get; set; }
        public long? ReorganizedDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
