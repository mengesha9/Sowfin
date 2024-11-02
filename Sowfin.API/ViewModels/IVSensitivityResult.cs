using Sowfin.API.ViewModels.InternalValuation;

namespace Sowfin.API.ViewModels
{
    public class IVSensitivityResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<KeyValueViewModel> Result { get; set; }
        }
}

