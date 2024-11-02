using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.FAnalysis
{
    public class FinancialStatementAnalysisDatasViewModel
    {
        public long Id { get; set; }
        public long? InitialSetup_FAnalysisId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public long? IntegratedDatasFAnalysisId { get; set; }
        public List<FinancialStatementAnalysisValuesViewModel> FinancialStatementAnalysisValuesVM { get; set; }
    }
    public class FinancialStatementAnalysisValuesViewModel
    {
        public long Id { get; set; }
        public long? FinancialStatementAnalysisDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
    public class FinancialStatementAnalysisFilingsViewModel
    {       
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public string CIK { get; set; }
        public List<FinancialStatementAnalysisDatasViewModel> FinancialStatementAnalysisDatasVM { get; set; }
    }
}
