using System;
using System.Collections.Generic;

namespace Sowfin.Model.Entities
{
    public class ROICDatas
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public bool?  IsParentItem { get; set; }
        public bool IsExplicit_editable { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<ROICValues> ROICValues { get; set; } = new List<ROICValues>();
        public List<ROIC_ExplicitValues> ROIC_ExplicitValues { get; set; } = new List<ROIC_ExplicitValues>();
        public string DtValue { get; set; }
    }

    public class ROICValues
    {
        public long Id { get; set; }
        public long? ROICDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }

    public class ROIC_ExplicitValues
    {
        public long Id { get; set; }
        public long? ROICDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
