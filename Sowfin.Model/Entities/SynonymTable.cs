using System;
using System.Collections.Generic;

namespace Sowfin.Model
{
    public class SynonymTable
    {
  
        public long Id { get; set; }
        public string FinField { get; set; }
        public string OtherTags { get; set; }
        public string Category { get; set; }
        public string Synonym { get; set; }
        public string StatementType { get; set; }
    }
}