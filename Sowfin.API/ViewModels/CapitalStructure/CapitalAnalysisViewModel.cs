using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalStructure
{
    public class CapitalAnalysisViewModel
    {
        public long Id { get; set; }
        public string LeveragePolicy { get; set; }
        public string AnalysisObject { get; set; }
        public long UserId { get; set; }
        public string SummaryOutput { get; set; }
        public string PermanentDebtUnit { get; set; }
        public string FreeCashFlowUnit { get; set; }
    }
}
