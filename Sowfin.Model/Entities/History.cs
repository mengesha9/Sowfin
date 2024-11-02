using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public partial class History
    {
        public long Id { get; set; }
        public long UserId { get; set; }
        public string PayoutTable { get; set; }
        public string ShareTable { get; set; }
        public string CurrentCapitalTable { get; set; }
        public string CurrentCostOfCapTable { get; set; }
        public string OtherInputTable { get; set; }
        public string FinancialTable { get; set; }
        public string SummaryOutput { get; set; }
        public int ApprovalFlag { get; set; }
        public int StartYear { get; set; }
        public int EndYear { get; set; }
        public int SummaryFlag { get; set; }

    }
}
