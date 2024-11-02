using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalStructure
{
    public class CapitalAnalysisSnapshots
    {
        public long Id { get; set; }
        public string SnapShot { get; set; }
        public string Description { get; set; }
        public long UserId { get; set; }
    }

    public class CapitalStructureScenarioSnapshotsViewModel
    {
        public long Id { get; set; }
        public long MasterId { get; set; }
        public string Description { get; set; }
        public int UserId { get; set; }
        public object scenarioOutput { get; set; }
        public object scenarioPieChart { get; set; }
        public int? LeveragePolicyId { get; set; }
        public DateTime? CreatedAt { get; set; }
        public bool Active { get; set; }
        public string SnapshotType { get; set; }
    }
}
