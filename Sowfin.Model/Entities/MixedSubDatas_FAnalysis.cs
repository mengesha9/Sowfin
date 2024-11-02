using System;
using System.Collections.Generic;
using System.Text;


namespace Sowfin.Model.Entities
{
    public class MixedSubDatas_FAnalysis
    {
        public long Id { get; set; }
        public long DatasId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public long? InitialSetup_FAnalysisId { get; set; }
        public List<MixedSubValues_FAnalysis> MixedSubValues_FAnalysis { get; set; } 
    }

    public class MixedSubValues_FAnalysis
    {
        public long Id { get; set; }
        public long MixedSubDatas_FAnalysisId { get; set; }
        public string FilingDate { get; set; }
        public string Value { get; set; }
    }
}
