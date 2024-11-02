using Sowfin.API.ViewModels.InternalValuation;

namespace Sowfin.API.ViewModels
{


       public class IntegratedResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public string AdjustedMessage { get; set; }
            public bool DepreciationFlag { get; set; }
            public List<IntegratedFilingsViewModel> Result { get; set; }
        }

}


