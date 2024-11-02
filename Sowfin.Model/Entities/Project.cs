using System;

namespace Sowfin.Model.Entities
{
    public class Project
    {
        public long Id { get; set; }
        public long BusinessUnitId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int? ValuationTechniqueId { get; set; }
        public int? StartingYear { get; set; }
        public int? NoOfYears { get; set; }
        public int? UserId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public bool?  IsActive { get; set; }
        public int? ApprovalFlag { get; set; }
        public string ApprovalComment { get; set; }
    }

    // old one
    public class Projects
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public long BusinessUnitId { get; set; }
        public long ValuationMethodId { get; set; }
        public string BusinessUnitName { get; set; }
        public string ValuationMethodName { get; set; }
    }
}
