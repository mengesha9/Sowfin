using System;
using System.Collections.Generic;

namespace Sowfin.Model.Entities
{
    public class EdgarView
    {
        public EdgarView()
        {
        }

        public long Id { get; set; }
        public string CIK { get; set; }
        public string StatementType { get; set; }
        public string LineItem { get; set; }
    }

    public class EdgarViewByCategory
    {
        public EdgarViewByCategory()
        {
        }

        public long Id { get; set; }
        public string CIK { get; set; }
        public string StatementType { get; set; }
        public string LineItem { get; set; }
    }
}
