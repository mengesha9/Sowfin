using Sowfin.API.ViewModels.InternalValuation;
namespace Sowfin.API.ViewModels
{

 public class ValuationResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<ValuationFilingsViewModel> Result { get; set; }
        }
}

    