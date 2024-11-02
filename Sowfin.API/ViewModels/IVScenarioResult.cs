using Sowfin.API.ViewModels.InternalValuation;

namespace Sowfin.API.ViewModels
{

        public class IVScenarioResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<IVScenarioViewModel> Result { get; set; }
        }
}

  