using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public partial class CapitalBudgeting
    {
        public long Id { get; set; }
        public int StartingYear { get; set; }
        public int NoOfYears { get; set; }
        public double DiscountRate { get; set; }
        public double MarginalTaxRate { get; set; }
        public string TableData { get; set; }
        public long UserId { get; set; }
        public int ApprovalFlag { get; set; }
        public string SummaryOutput { get; set; }
        public string RawTableData { get; set; }
        public long ProjectId { get; set; }
        public int RevenueCount { get; set; }
        public int CapexCount { get; set; }
        public int FixedCostCount { get; set; }
        public string NPV { get; set; }
        public string ApprovalComment { get; set; }
    }
}
