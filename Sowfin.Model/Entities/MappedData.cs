using System;


namespace Sowfin.Model.Entities
{
    public class MappedData
    {
        public long Id { get; set; }
        public string FilingDate { get; set; }
        public string Cik { get; set; }
        public string LineItem { get; set; }
        public string ParentItem { get; set; }
        public string Value { get; set; }
        public string StatementType { get; set; }
        public string Category { get; set; }
        public string OtherTags { get; set; }
        public int Sequence { get; set; }
    }
}
