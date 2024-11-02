using Sowfin.API.ViewModels.PayoutPolicy;
using  Sowfin.Model.Entities;
using Sowfin.Data.Common.Helper;
namespace Sowfin.API.ViewModels
{
    public class CurrentSetup_PayoutPolicyResult
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public bool IsSaved { get; set; }
        public CurrentSetupViewModel currentSetupObj { get; set; }
        public List<InitialSetup_IValuation> InitialSetup_IValuationList { get; set; }
        public List<CurrentSetupIpFillingsViewModel> FilingResult { get; set; }
        public List<ValueTextWrapper> currencyValueList { get; set; } //Date:23-July-2020 | Added By : anonymous | Enh. : Single Units Enum 
        public List<ValueTextWrapper> numberCountList { get; set; } //Date:23-July-2020 | Added By : anonymous | Enh. : Single Units Enum 
    }
}
