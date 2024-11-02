using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.History
{
    public class HistoryObject
    {
        public long id { get; set; }
        public long userId { get; set; }
        public object[][] payoutTable { get; set; }
        public object[][] shareTable { get; set; }
        public object[][] currentCapitalTable { get; set; }
        public object[][] currentCostOfCapTable { get; set; }
        public object[][] otherInputTable { get; set; }
        public object[][] financialTable { get; set; }
        public int startYear { get; set; }
        public int endYear { get; set; }
        public string summaryOutput { get; set; }
        public int summaryFlag { get; set; }

    }
}
