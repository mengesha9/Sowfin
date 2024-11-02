using System.Collections.Generic;

namespace Sowfin.API.ViewModels
{
    public class UpdateStatementViewModel
    {
        public string LineItem { get; set; }
        public string ParentItem { get; set; }
        public string Category { get; set; }
        public string OtherTags { get; set; }
        public bool IsMultiInstances { get; set; }
        public string Description { get; set; }
        public int Sequence { get; set; }
        public List<string> Synonyms { get; set; } = new List<string>();
    }
}