using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class ForcastRatioFilingsViewModel
    {
        //public long Id { get; set; }
        //public string CIK { get; set; }
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public string CIK { get; set; }

        public List<ForcastRatioDatasViewModel> ForcastRatioDatasVM { get; set; }
    }
    public class ForcastRatioDatasViewModel
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
        public List<ForcastRatioValuesViewModel> ForcastRatioValuesVM { get; set; }
        public List<ForcastRatio_ExplicitValuesViewModel> ForcastRatio_ExplicitValuesVM { get; set; }
        public long? ParentId { get; set; }
        public string ChildId { get; set; }
        public bool? isParent { get; set; }
    }
    public class ForcastRatioValuesViewModel
    {
        public long Id { get; set; }
        public long? ForcastRatioDatasId { get; set; }
        public string FilingDate { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }

    public class ForcastRatio_ExplicitValuesViewModel
    {
        public long Id { get; set; }
        public long? ForcastRatioDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
