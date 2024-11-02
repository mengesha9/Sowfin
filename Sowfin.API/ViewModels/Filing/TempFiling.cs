using Newtonsoft.Json;
using System;

namespace Sowfin.API.ViewModels.Filing
{
    public class TempFiling
    {
        [JsonProperty(PropertyName = "id")]
        public long Id { get; set; }

        [JsonProperty(PropertyName = "category")]
        public string Category { get; set; }

        [JsonProperty(PropertyName = "lineItem")]
        public string LineItem { get; set; }

    }
}
