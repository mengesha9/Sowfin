using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System;

namespace Sowfin.API.ViewModels
{
    public class FindataViewModel
    {
        public long id { get; set; }
        public string cik { get; set; }
        public string filingDate { get; set; }
        public string parentItem { get; set; }
        public string lineItem { get; set; }
        public string value { get; set; }
        public string statementType { get; set; }
        public string category { get; set; }
        public string otherTags { get; set; }
        public string sequence { get; set; }
        public string company { get; set; }
        public string finField { get; set; }


    }

    public class FinCell
    {
        public string LineItem { get; set; }
        public string FilingDate { get; set; }
        public string Value { get; set; }
    }

    public class FinRow
    {
        public string LineItem { get; set; }
        public List<FinCell> FinCells { get; set; }
    }

    public class LineItemEx
    {
        public string LineItem { get; set; }
    }

    public class FindataViewModelEx
    {
        public long Id { get; set; }

        public string Cik { get; set; }

        public string FilingDate { get; set; }

        public string ParentItem { get; set; }

        public string LineItem { get; set; }

        public string StatementLineItem { get; set; }

        public string Value { get; set; }

        public string StatementType { get; set; }

        public string Category { get; set; }

        public string OtherTags { get; set; }

        public int Sequence { get; set; }
    }

}