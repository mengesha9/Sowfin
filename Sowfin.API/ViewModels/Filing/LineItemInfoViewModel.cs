using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace Sowfin.Model.Entities
{
    public class LineItemInfoViewModel
    {
       

        [JsonProperty(PropertyName = "id")]
        public long Id { get; set; }


        [JsonProperty(PropertyName = "category")]
        public string Category { get; set; }

        public List<MixedSubDatas> MixedSubDatas { get; set; }
        public List<MixedSubDatas_FAnalysis> MixedSubDatas_FAnalysis { get; set; }

    }




}

