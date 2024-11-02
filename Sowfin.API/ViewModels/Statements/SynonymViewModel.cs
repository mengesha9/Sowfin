using Sowfin.Model;
using System;
using System.Collections.Generic;

namespace Sowfin.API.ViewModels
{
    public class SynonymViewModel
    {
        public long SynonymId { get; set; }
        public long StatementId { get; set; }
        public string Word { get; set; }
    }
}