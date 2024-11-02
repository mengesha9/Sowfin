using System;
using System.Collections.Generic;
using System.Text;
namespace Sowfin.Model.Entities
{
    public class MasterCostofCapitalNStructure
    {
        public long Id { get; set; }
        public int? LeveragePolicyId { get; set; }
        public int? SALeveragePolicyID { get; set; }
        public bool HasEquity { get; set; }
        public bool HasPreferredEquity { get; set; }
        public bool HasDebt { get; set; }
        public string Company { get; set; }
        public int? BetaSourceId { get; set; }
        public int CostofDebtMethodId { get; set; }
        public long? UserId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }

    public class Snapshot_CostofCapitalNStructure
    {
        public long Id { get; set; }
        public int? MasterId { get; set; }
        public string Description { get; set; }
        public string Comment { get; set; }
        public bool Active { get; set; }
        public DateTime? CreatedDate { get; set; }
        // public List<CostofCapital_Snapshot> CostofCapitalSnapshotList { get; set; } 
        // public List<CapitalStructure_Snapshot> CapitalStructureSnapshotList { get; set; } 
    }

    public class CostofCapital_Snapshot
    {
        public long Id { get; set; }
        public long? Snapshot_CostofCapitalNStructureId { get; set; }
        public string LineItem { get; set; }
        public int HeaderId { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }
    }

    public class CapitalStructure_Snapshot
    {
        public long Id { get; set; }
        public long? Snapshot_CostofCapitalNStructureId { get; set; }
        public string LineItem { get; set; }
        public int HeaderId { get; set; }
        public int? DefaultUnitId { get; set; }
        public int? UnitId { get; set; }
        public int ValueTypeId { get; set; }
        public double? BasicValue { get; set; }
        public double? Value { get; set; }
    }
}
