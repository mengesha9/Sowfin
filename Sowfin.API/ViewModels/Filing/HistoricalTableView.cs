using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.Filing
{
    public class HistoricalTableView
    {
        [JsonProperty("id")]
        public long Id { get; set; }

        [JsonProperty("statementType")]
        public string StatementType { get; set; }

        [JsonProperty("lineItem")]
        public string LineItem { get; set; }

        [JsonProperty("sequence")]
        public int Sequence { get; set; }

        [JsonProperty("category")]
        public string Category { get; set; }

        [JsonProperty("finField")]
        public string FinField { get; set; }


        [JsonProperty("cik")]
        public string Cik { get; set; }

        [JsonProperty("years")]
        public List<string> Years { get; set; } = new List<string>();

        [JsonProperty("values")]
        public List<string> Values { get; set; } = new List<string>();

        [JsonProperty("isParent")]
        public bool IsParent { get; set; }

        [JsonProperty("ParentItem")]
        public string ParentItem { get; set; }
    }
}
