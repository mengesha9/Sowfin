using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class ProjectInputDatas
    {
        public long Id { get; set; }
        public long? ProjectId { get; set; }
        public long? HeaderId { get; set; }
        public string SubHeader { get; set; }
        public string LineItem { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int? ValueTypeId { get; set; }
        public bool HasMultiYear { get; set; }
        public double? Value { get; set; }
        public List<ProjectInputValues> ProjectInputValues { get; set; } 
        public List<ProjectInputComparables> ProjectInputComparables { get; set; } 
        public List<DepreciationInputValues> DepreciationInputValues { get; set; } 
    }

    public class ProjectInputValues
    {
        public long Id { get; set; }
        public long? ProjectInputDatasId { get; set; }
        public int? Year { get; set; }
        public double? Value { get; set; }
    }

    public class ProjectInputComparables
    {
        public long Id { get; set; }
        public long? ProjectInputDatasId { get; set; }
        public string Comparable { get; set; }
        public int? Count { get; set; }
        public double? Value { get; set; }
    }

    public class DepreciationInputValues
    {
        public long Id { get; set; }
        public long? ProjectInputDatasId { get; set; }
        public int? Year { get; set; }
        public double? Value { get; set; }
    }
}
