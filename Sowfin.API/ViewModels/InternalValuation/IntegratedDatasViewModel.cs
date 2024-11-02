using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class IntegratedDatasViewModel
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
        public List<IntegratedValuesViewModel> IntegratedValuesVM { get; set; }
        public List<Integrated_ExplicitValuesViewModel> Integrated_ExplicitValuesVM { get; set; }
    }
    public class IntegratedValuesViewModel
    {
        public long Id { get; set; }
        public long? IntegratedDatasId { get; set; }
        public string FilingDate { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }


    public class Integrated_ExplicitValuesViewModel
    {
        public long Id { get; set; }
        public long? IntegratedDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }

    public class IntegratedFilingsViewModel
    {
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public string CIK { get; set; }
        public List<IntegratedDatasViewModel> IntegratedDatasVM { get; set; }
    }
}
