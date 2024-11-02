using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
  //public class ScenarioInputData
  //  {
  //      public long Id { get; set; }
  //      //public long? ProjectId { get; set; }
  //      //public string Name { get; set; }
  //      //public string Description { get; set; }
  //      //public double? Probability { get; set; }
  //      //public string LineItem { get; set; }
  //      //public double? NPV { get; set; }
  //      public List<ScenarioInputDatas> ScenarioInputDatas { get; set; }       
  //  }
   public class ScenarioInputDatas
    {
        public long Id { get; set; }
        public long? ScenarioInputDatasId { get; set; }
        public long? ProjectId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public double? Probability { get; set; }
        // public string LineItem { get; set; }
        public double? NPV { get; set; }
        public List<ScenarioInputValues> ScenarioInputValues { get; set; } = new List<ScenarioInputValues>();
    }

    public class ScenarioInputValues
    {
        public long Id { get; set; }
        public long? ScenarioInputDatasId { get; set; }
        public string LineItem { get; set; }
        public double? LineItemValue { get; set; }
    }
}
