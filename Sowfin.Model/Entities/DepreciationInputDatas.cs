using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class DepreciationInputDatas 
  {
      public long Id { get; set; }
      public long? ProjectId { get; set; }
      public long? ProjectInputDatasId { get; set; }
      // public string LineItem { get; set; }
      public bool HasMultiYear { get; set; }
      public long? Method { get; set; }
      // public long? DepreciationMethodValue { get; set; }
      public bool SameYear { get; set; }
      public long? Duration { get; set; }
      // public List<DepreciationInputValues> DepreciationInputValues { get; set; }
  }

    //public class DepreciationInputValues
    //{
    //    public long Id { get; set; }
    //    public long? DepreciationInputDatasId { get; set; }
    //    public int? Year { get; set; }
    //    public double? Value { get; set; }

    //}

}
