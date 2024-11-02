using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
     public class HistoryElementMapping
    {
        public long Id { get; set; }
        public string ItemName { get; set; }
        public string CommonElementName { get; set; }
        public int StatementTypeId { get; set; }
        public string ItemCode { get; set; }
    }
}
