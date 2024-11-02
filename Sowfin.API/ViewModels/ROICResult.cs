using Sowfin.API.ViewModels.InternalValuation;

namespace Sowfin.API.ViewModels
{
        public class ROICResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<ROICFilingsViewModel> Result { get; set; }
        }
}

