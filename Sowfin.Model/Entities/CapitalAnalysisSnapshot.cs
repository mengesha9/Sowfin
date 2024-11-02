using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class CapitalAnalysisSnapshot
    {
      
        public long Id { get; set; }
        public string SnapShot { get; set; }
        public string Description { get; set; }
        public long UserId { get; set; }
    }
}
