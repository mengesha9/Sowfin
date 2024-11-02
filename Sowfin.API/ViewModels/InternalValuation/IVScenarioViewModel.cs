using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class IVScenarioViewModel
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string Scenario { get; set; }
        public string Probability { get; set; }
        public string WACC { get; set; }
        public string TGR { get; set; }
        public string NetPPE { get; set; }
        public string StockPrice { get; set; }
    }

    public class IVScenarioListViewModel
    {
        public List<IVScenarioViewModel> IVScenarioListVM { get; set; }
    }
}
