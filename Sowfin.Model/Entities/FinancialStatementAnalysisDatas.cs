using System;
using System.Collections.Generic;
using System.Text;


namespace Sowfin.Model.Entities
{
    public class FinancialStatementAnalysisDatas
    {
        public long Id { get; set; }
        public long? InitialSetup_FAnalysisId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public long? IntegratedDatasFAnalysisId { get; set; }
        public List<FinancialStatementAnalysisValues> FinancialStatementAnalysisValues { get; set; } = new List<FinancialStatementAnalysisValues>();
    }
    
    public class FinancialStatementAnalysisValues
    {
        public long Id { get; set; }
        public long? FinancialStatementAnalysisDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}

