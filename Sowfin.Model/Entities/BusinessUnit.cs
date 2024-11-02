using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class BusinessUnit
    {
 
        public long Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int Status { get; set; }
        public int YearEstablished { get; set; }
        public long UserId { get; set; }
    }
}
