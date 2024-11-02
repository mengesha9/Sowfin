using Sowfin.API.ViewModels.InternalValuation;
namespace Sowfin.API.ViewModels
{
        public class AssetsEquityResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<AssetsEquityFilingsViewModel> Result { get; set; }
        }

}
