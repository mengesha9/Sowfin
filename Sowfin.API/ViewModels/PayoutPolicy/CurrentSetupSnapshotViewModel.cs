using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.PayoutPolicy
{
    public class CurrentSetupSnapshotFillingsViewModel
    {
        public string StatementType { get; set; }
        public List<CurrentSetupSnapshotDatasViewModel> CurrentSetupSnapshotDatasVM { get; set; }
    }
    public class CurrentSetupSnapshotViewModel
    {
        public long Id { get; set; }
        public string Description { get; set; }
        public long UserId { get; set; }
        public List<CurrentSetupSnapshotDatasViewModel> CurrentSetupSnapshotDatasVM { get; set; }
        public List<CurrentSetupSoFillingsViewModel> CurrentSetupSoFillingsVM { get; set; }
    }
    public class CurrentSetupSnapshotDatasViewModel
    {
        public long Id { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public long Sequence { get; set; }
        public int StatementTypeId { get; set; }
        public string Unit { get; set; }
        public long? CurrentSetupSnapshotId { get; set; }
        public List<CurrentSetupSnapshotValuesViewModel> CurrentSetupSnapshotValuesVM { get; set; }
    }
    public class CurrentSetupSnapshotValuesViewModel
    {
        public long Id { get; set; }
        public long? CurrentSetupSnapshotDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
