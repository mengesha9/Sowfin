using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class CapitalStructureScenarioSnapshot
    {
        public long Id { get; set; }
        public long? MasterId { get; set; }
        public string? Description { get; set; } 
        public int UserId { get; set; }
        public string? scenarioOutput { get; set; } 
        public string? scenarioPieChart { get; set; } 
        public int? LeveragePolicyId { get; set; }
        public DateTime? CreatedAt { get; set; } 
        public bool Active { get; set; }
        public string? SnapshotType { get; set; } 
    }
}
    
