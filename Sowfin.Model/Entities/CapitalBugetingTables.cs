using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class CapitalBugetingTables
    {
   
        public long Id { get; set; }
        public long ProjectId { get; set; }
        public long UserId { get; set; }
        public int StartingYear { get; set; }
        public int NoOfYears { get; set; }
        public double DiscountRate { get; set; }
        public double MarginalTax { get; set; }
        public double EvalFlag { get; set; }
        public double ApprovalFlag { get; set; }
        public string Tables { get; set; }
    }
}
