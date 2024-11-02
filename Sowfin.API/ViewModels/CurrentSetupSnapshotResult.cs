using Sowfin.API.ViewModels.PayoutPolicy;
namespace Sowfin.API.ViewModels
{
    public class CurrentSetupSnapshotResult
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public List<CurrentSetupSnapshotFillingsViewModel> Result { get; set; }
    }
}

