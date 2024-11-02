using System;
using System.Collections.Generic;
using System.Text;
namespace Sowfin.Model.Entities
{
    public class MarketDatas
    {
        public long Id { get; set; }
        public long? InitialSetup_FAnalysisId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public string? Comments { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<MarketValues> MarketValues { get; set; } = new List<MarketValues>();
    }

    public class MarketValues
    {
        public long Id { get; set; }
        public long? MarketDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}