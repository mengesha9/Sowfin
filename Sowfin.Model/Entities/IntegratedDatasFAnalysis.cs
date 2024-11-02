using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class IntegratedDatasFAnalysis
    {
        public long Id { get; set; }
        public long? InitialSetup_FAnalysisId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public bool?  IsParentItem { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<IntegratedValuesFAnalysis> IntegratedValuesFAnalysis { get; set; } = new List<IntegratedValuesFAnalysis>();
    }

    public class IntegratedValuesFAnalysis
    {
        public long Id { get; set; }
        public long? IntegratedDatasFAnalysisId { get; set; }
        public string FilingDate { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}

