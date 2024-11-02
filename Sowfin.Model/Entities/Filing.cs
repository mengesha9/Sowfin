using System;

namespace Sowfin.Model
{
    public class Filing
    {

        public long Id { get; set; }
        public string? FilingDate { get; set; }
        public string? Cik { get; set; }
        public string? LineItem { get; set; }
        public string? Company { get; set; }
        public string? Value { get; set; }
        public string? ParentItem { get; set; }
        public string? StatementType { get; set; }
        public int Sequence { get; set; }
    }
}