using System;
using System.Collections.Generic;
using System.Text;
namespace Sowfin.Model.Entities
{
    public class MixedData
    {
        public long Id { get; set; }
        public long DatasId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public long? InitialSetupId { get; set; }
        public List<MixedValues> MixedValues { get; set; }
    }

    public class MixedValues
    {
        public long Id { get; set; }
        public long MixedSubDatasId { get; set; }
        public string FilingDate { get; set; }
        public string Value { get; set; }
    }
}
