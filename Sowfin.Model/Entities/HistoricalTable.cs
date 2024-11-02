using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class HistoricalTable
    {

        public long Id { get; set; }
        public string Cik { get; set; }
        public string LineItem { get; set; }
        public int Sequence { get; set; }
        public string StatementType { get; set; }
        public string Category { get; set; }
        public string FinField { get; set; }
        public List<string> Years { get; set; } = new List<string>();
        public List<string> Values { get; set; } = new List<string>();
        public bool IsParent { get; set; }
        public string ParentItem { get; set; }


    }
}
