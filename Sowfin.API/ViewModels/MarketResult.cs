using Sowfin.API.ViewModels.FAnalysis;

namespace Sowfin.API.ViewModels{

        public class MarketResult
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
            public List<MarketDatasViewModel> Result { get; set; }
        }

}
