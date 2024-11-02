using System;
using System.Collections.Generic;


namespace Sowfin.Model.Entities
{
       public class ProcessedData
    {

        public long Id { get; set; }
        public string? Cik { get; set; } 
        public string LineItem { get; set; } = string.Empty;
        public int Sequence { get; set; }
        public string StatementType { get; set; } = string.Empty; 
        public string Category { get; set; } = string.Empty; 
        public string? FinField { get; set; } 
        public List<string> Years { get; set; } = new List<string>(); 
        public List<string> Values { get; set; } = new List<string>();
        public bool IsParent { get; set; }
        public string? ParentItem { get; set; } 
    }
}



