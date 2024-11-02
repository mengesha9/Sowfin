using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class ReorganizedFilingsViewModel
    {
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public string CIK { get; set; }

        public List<ReorganizedDatasViewModel> ReorganizedDatasVM { get; set; }
    }
    public class ReorganizedDatasViewModel
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public bool? IsParentItem { get; set; }
        public long? IntegratedDatasId { get; set; }
        public bool IsExplicit_editable { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<ReorganizedValuesViewModel> ReorganizedValuesVM { get; set; }
        public List<Reorganized_ExplicitValuesViewModel> Reorganized_ExplicitValuesVM { get; set; }
    }
    public class ReorganizedValuesViewModel
    {
        public long Id { get; set; }
        public long? ReorganizedDatasId { get; set; }
        public string FilingDate { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
    public class Reorganized_ExplicitValuesViewModel
    {
        public long Id { get; set; }
        public long? ReorganizedDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
