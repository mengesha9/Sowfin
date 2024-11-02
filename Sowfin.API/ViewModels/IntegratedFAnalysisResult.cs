using Sowfin.API.ViewModels.FAnalysis;
namespace Sowfin.API.ViewModels{
        public class IntegratedFAnalysisResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public string AdjustedMessage { get; set; }
            public bool DepreciationFlag { get; set; }
            public List<IntegratedFilingsFAnalysisViewModel> Result { get; set; }
        }
}
