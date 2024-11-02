using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
   public class CIKStatus
    {
        public long Id { get; set; }
        public String CIK { get; set; }
        public Int16 Status { get; set; }
        public String Remark { get; set; }

    }
}
