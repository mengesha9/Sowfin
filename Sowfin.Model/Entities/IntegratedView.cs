using System;
using System.Collections.Generic;
using System.Text;



namespace Sowfin.Model.Entities
{
    public class IntegratedView
    {
    }

    public class FilingsTable
    {
        public long Id { get; set; }
        public string CIK { get; set; }
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public long? InitialSetupId { get; set; }
        public short? Sequence { get; set; }
        //  public List<DatasTable> Datas { get; set; }
    }

    public class DatasTable
    {
        public long Id { get; set; }
        public long FilingId { get; set; }
        public string ElementName { get; set; }
        public string ParentItem { get; set; }
        public string LineItem { get; set; }
        public bool IsParentItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public bool?  IsSegmentTitle { get; set; }
        public long InitialSetupId { get; set; }
        public List<ValuesTable> Values { get; set; } = new List<ValuesTable>();
        public List<MixedSubDatas> MixedSubDatas { get; set; } = new List<MixedSubDatas>();
    }

    public class ValuesTable
    {
        public long Id { get; set; }
        public long DataId { get; set; }
        public string FilingDate { get; set; }
        public string Value { get; set; }
        public string CElementName { get; set; }
        public string CLineItem { get; set; }
    }

    public class IntegratedFilingsArray
    {
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public IntegratedFilings Filings { get; set; }
    }

    public class IntegratedFilings
    {
        //public long Id { get; set; }
        //public string CIK { get; set; }
        public string CompanyName { get; set; }
        public string StatementType { get; set; }
        public string ReportName { get; set; }
        public string Unit { get; set; }
        public string CIK { get; set; }
        public List<IntegratedDatas> IntegratedDatas { get; set; } = new List<IntegratedDatas>();
    }

    public class IntegratedDatas
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string LineItem { get; set; }
        public bool IsTally { get; set; }
        public long Sequence { get; set; }
        public string Category { get; set; }
        public int StatementTypeId { get; set; }
        public bool?  IsParentItem { get; set; }
        public bool IsExplicit_editable { get; set; }
        public bool IsHistorical_editable { get; set; }
        public List<IntegratedValues> IntegratedValues { get; set; } = new List<IntegratedValues>();
        public List<Integrated_ExplicitValues> integrated_ExplicitValues { get; set; } = new List<Integrated_ExplicitValues>();
    }

    public class IntegratedValues
    {
        public long Id { get; set; }
        public long? IntegratedDatasId { get; set; }
        public string FilingDate { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }

    public class Integrated_ExplicitValues
    {
        public long Id { get; set; }
        public long? IntegratedDatasId { get; set; }
        public string Year { get; set; }
        public string Value { get; set; }
    }
}
