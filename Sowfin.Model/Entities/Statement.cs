using System;
using System.Collections.Generic;

namespace Sowfin.Model
{
    public class Statement
    {

        public long StatementId { get; set; }
        public string LineItem { get; set; }
        public string Category { get; set; }
        public string OtherTags { get; set; }
        public bool IsMultiInstances { get; set; }
        public string Description { get; set; }
        public int Sequence { get; set; }
        public string StatementType { get; set; }
        public List<string> Synonyms { get; set; } = new List<string>();
    }
}