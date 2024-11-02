using System;
using System.Collections.Generic;
using System.Text;


namespace Sowfin.Model.Entities
{
    public class ForcastRatioDatas
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public bool?  IsParentItem { get; set; }
        public long? IntegratedDatasId { get; set; }
        public bool IsExplicit_editable { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<ForcastRatioValues> ForcastRatioValues { get; set; } = new List<ForcastRatioValues>();
        public List<ForcastRatio_ExplicitValues> ForcastRatio_ExplicitValues { get; set; } = new List<ForcastRatio_ExplicitValues>();
        public long? ParentId { get; set; }
        public string ChildId { get; set; }
        public bool?  isParent { get; set; }
    }

    public class ForcastRatioValues
    {
        public long Id { get; set; }
        public long? ForcastRatioDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
    
    public class ForcastRatio_ExplicitValues
    {
        public long Id { get; set; }
        public long? ForcastRatioDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}

