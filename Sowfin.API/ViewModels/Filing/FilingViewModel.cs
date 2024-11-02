using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System;

namespace Sowfin.API.ViewModels
{
    public class FilingViewModel
    {
        public long Id { get; set; }

        public string Cik { get; set; }

        public string FilingDate { get; set; }

        public string ParentItem { get; set; }
            
        public string LineItem { get; set; }

        public string Value { get; set; }

        public string StatementType { get; set; }
    }
    
}