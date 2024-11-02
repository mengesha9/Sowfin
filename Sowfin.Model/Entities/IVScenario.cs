using System;
using System.Collections.Generic;
using System.Text;


namespace Sowfin.Model.Entities
{
    public class IVScenario
    {
        public long Id { get; set; }
        public long? InitialSetupId { get; set; }
        public string Scenario { get; set; }
        public string Probability { get; set; }
        public string WACC { get; set; }
        public string TGR { get; set; }
        public string NetPPE { get; set; }
        public string StockPrice { get; set; }
    }
}
