using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class Snapshot_CapitalBudgeting
    {
        public int Id { get; set; }
        public string ProjectInputs { get; set; }
        public string Snapshot { get; set; }
        public string Description { get; set; }
        public int UserId { get; set; }
        public int ProjectId { get; set; } 
        public int ValuationTechniqueId { get; set; }
        public string NPV { get; set; }
        public string CNPV { get; set; }
        public DateTime CreatedAt { get; set; }
        public bool Active { get; set; }
        public string SnapshotType { get; set; }
    }
}
