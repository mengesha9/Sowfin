using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class Approval 
    {
        public long ProjectId { get; set; }
        public int ApprovalFlag { get; set; }
        public string ApprovalPassword { get; set; }
        public string ApprovalComment { get; set; }
    }
}
