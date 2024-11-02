using Sowfin.API.ViewModels.PayoutPolicy;
namespace Sowfin.API.ViewModels
{
    public class PayoutPolicy_ScenarioResult
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public string ReportName { get; set; }
        public bool IsSaved { get; set; }
        public List<PayoutPolicy_ScenarioFillingsViewModel> FilingResult { get; set; }
    }
}