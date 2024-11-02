using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class IntegratedFinancialStmt
    {
        public long Id { get; set; }
        public int? IntegratedElementMasterId { get; set; }
        public long? InitialSetupId { get; set; }
        public int Year { get; set; }
        public string Value { get; set; }
    }
}
