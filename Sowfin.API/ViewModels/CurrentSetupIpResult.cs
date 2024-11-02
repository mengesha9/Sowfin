using Sowfin.API.ViewModels.PayoutPolicy;
namespace Sowfin.API.ViewModels
{
    public class CurrentSetupIpResult
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public List<CurrentSetupIpFillingsViewModel> Result { get; set; }
    }
}
