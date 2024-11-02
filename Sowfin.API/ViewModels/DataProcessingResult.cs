using Sowfin.Model.Entities;

namespace Sowfin.API.ViewModels
{
    public class DataProcessingResult
        {
            public int StatusCode { get; set; }
            public long? InitialSetupId { get; set; }
            public List<FAnalysis_CategoryByInitialSetup> categoryList { get; set; }
            public List<FilingsArray> Result { get; set; }
            public int ShowMsg { get; set; }
        }
}
   