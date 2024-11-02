using Sowfin.API.ViewModels.FAnalysis;

namespace Sowfin.API.ViewModels
{
       public class FinancialStatementAnalysisResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<FinancialStatementAnalysisFilingsViewModel> Result { get; set; }
        }
}
