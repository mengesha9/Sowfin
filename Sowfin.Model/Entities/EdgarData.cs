using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class EdgarData
    {
        public string EdgarView { get; set; }
    }

    public class EdgarDataByCategory
    {
        public string EdgarViewByCategory { get; set; }
    }

    public class RenderResult
    {
        public int StatusCode { get; set; }
        public long? InitialSetupId { get; set; }        
        public List<CategoryByInitialSetup> categoryList { get; set; } = new List<CategoryByInitialSetup>();
        public List<FilingsArray> Result { get; set; } = new List<FilingsArray>();
        public int ShowMsg { get; set; }
    }

    public class FilingsArray
    {
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public Filings Filings { get; set; }
    }

    public class Filings
    {
        //public long Id { get; set; }
        //public string CIK { get; set; }
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public List<Datas> Datas { get; set; } = new List<Datas>();
        public long? InitialSetupId { get; set; }
    }

    public class Datas
    {
        //public long Id { get; set; }
        public long DataId { get; set; }
        public long Sequence { get; set; }
        public string ParentItem { get; set; }
        public string ElementName { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public bool IsTally { get; set; }
        public string Category { get; set; }
        public bool?  IsSegmentTitle { get; set; }
        public long InitialSetupId { get; set; }
        public List<Values> Values { get; set; } = new List<Values>();
        public List<MixedSubDatas> MixedSubDatas { get; set; } = new List<MixedSubDatas>();
        public List<MixedSubDatas_FAnalysis> MixedSubDatas_FAnalysis { get; set; } = new List<MixedSubDatas_FAnalysis>();
    }

    public class Values
    {
        public string FilingDate { get; set; }
        public string Value { get; set; }
        public string CElementName { get; set; }
        public string CLineItem { get; set; }
    }
}
