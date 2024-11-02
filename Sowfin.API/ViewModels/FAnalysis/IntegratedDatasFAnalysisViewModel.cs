using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.FAnalysis
{
    public class IntegratedDatasFAnalysisViewModel
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
        public List<IntegratedValuesFAnalysisViewModel> IntegratedValuesFAnalysisVM { get; set; }
    }
    public class IntegratedValuesFAnalysisViewModel
    {
        public long Id { get; set; }
        public long? IntegratedDatasId { get; set; }
        public string FilingDate { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
    public class IntegratedFilingsFAnalysisViewModel
    {
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public string CIK { get; set; }
        public List<IntegratedDatasFAnalysisViewModel> IntegratedDatasFAnalysisVM { get; set; }
    }
}
