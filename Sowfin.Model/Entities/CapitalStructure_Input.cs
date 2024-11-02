using System;
using System.Collections.Generic;
using System.Text;
namespace Sowfin.Model.Entities
{
    public class CapitalStructure_Input
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public int HeaderId { get; set; }
        public string? SubHeader { get; set; }
        public string? LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }
    }

    public class CapitalStructure_Output
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public string? LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }
    }
}
