using Sowfin.API.ViewModels.PayoutPolicy;
namespace Sowfin.API.ViewModels
{
    public class PayoutPolicy_ScenarioOutputResult
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public string ReportName { get; set; }
        public List<PayoutPolicy_ScenarioOutputFillingsViewModel> FilingResult { get; set; }
    }
}