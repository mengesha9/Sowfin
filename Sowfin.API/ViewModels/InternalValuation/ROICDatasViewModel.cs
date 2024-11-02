using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class ROICFilingsViewModel
    {
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public string CIK { get; set; }

        public List<ROICDatasViewModel> ROICDatasVM { get; set; }
    }
    public class ROICDatasViewModel
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public bool?  IsParentItem { get; set; }
       // public long? IntegratedDatasId { get; set; }
        public bool IsExplicit_editable { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<ROICValuesViewModel> ROICValuesVM { get; set; }
        public List<ROIC_ExplicitValuesViewModel> ROIC_ExplicitValuesVM { get; set; }
        public string DtValue { get; set; }
    }
    public class ROICValuesViewModel
    {
        public long Id { get; set; }
        public long? ROICDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
    public class ROIC_ExplicitValuesViewModel
    {
        public long Id { get; set; }
        public long? ROICDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
