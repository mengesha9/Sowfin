using System;

namespace Sowfin.Model.Entities
{
    public class RawHistoricalValues
    {
        public long Id { get; set; }
        public long DataId { get; set; }
        public DateTime? FilingDate { get; set; }
        public int? Value { get; set; }
        public string? CElementName { get; set; }
        public string? CLineItem { get; set; }
    }
}
