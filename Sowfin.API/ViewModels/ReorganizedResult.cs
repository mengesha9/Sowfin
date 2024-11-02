using Sowfin.API.ViewModels.InternalValuation;
namespace Sowfin.API.ViewModels
{
       public class ReorganizedResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<ReorganizedFilingsViewModel> Result { get; set; }
        }

}
