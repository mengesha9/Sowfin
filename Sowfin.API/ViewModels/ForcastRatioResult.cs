using Sowfin.API.ViewModels.InternalValuation;
namespace Sowfin.API.ViewModels
{
           public class ForcastRatioResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<ForcastRatioFilingsViewModel> Result { get; set; }
        }

}
