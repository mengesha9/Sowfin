using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.FAnalysis
{
    public class MarketDatasViewModel
    {
        public long Id { get; set; }
        public long? InitialSetup_FAnalysisId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public string Comments { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<MarketValuesViewModel> MarketValuesVM { get; set; }
    }
    public class MarketValuesViewModel
    {
        public long Id { get; set; }
        public long? MarketDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
