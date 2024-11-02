using Newtonsoft.Json;


namespace Sowfin.API.ViewModels.Filing
{
    public class FilingAndSynonym
    {
        [JsonProperty(PropertyName = "id")]
        public long Id { get; set; }

        [JsonProperty(PropertyName = "finField")]
        public string FinField { get; set; }

        [JsonProperty(PropertyName = "category")]
        public string Category { get; set; }

        [JsonProperty(PropertyName = "statementType")]
        public string StatementType { get; set; }

        [JsonProperty(PropertyName = "lineItem")]
        public string LineItem { get; set; }

        [JsonProperty(PropertyName = "value")]
        public string Value { get; set; }

        [JsonProperty(PropertyName = "parentItem")]
        public string ParentItem { get; set; }

        [JsonProperty(PropertyName = "filingDate")]
        public string FilingDate { get; set; }

        [JsonProperty(PropertyName = "cik")]
        public string Cik { get; set; }

        [JsonProperty(PropertyName = "sequence")]
        public int Sequence { get; set; }
    }
}
