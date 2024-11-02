using System;
using System.ComponentModel.DataAnnotations;

namespace Sowfin.Model
{
    public class Findata
    {
        public Findata()
        {
        }

        [Key]
        public long Id { get; set; }
        public string FilingDate { get; set; }
        public string Cik { get; set; }
        public string LineItem { get; set; }
        public string Company { get; set; }
        public string Value { get; set; }
        public string ParentItem { get; set; }
        public string StatementType { get; set; }
        public string FinField { get; set; }
        public string Category { get; set; }
        public string OtherTags { get; set; }
        public int Sequence { get; set; }
    }
}
