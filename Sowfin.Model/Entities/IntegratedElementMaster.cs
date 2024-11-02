using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class IntegratedElementMaster
    {
        public long Id { get; set; }
        public string ElementName { get; set; }
        public int Sequence { get; set; }
        public bool IsFormulaReq { get; set; }
        public int? StatementTypeId { get; set; }
        public long? IntegratedReferenceId { get; set; }
        public long? HistoryElementMappingId { get; set; }
        public string ReferenceName { get; set; }
    }
}
