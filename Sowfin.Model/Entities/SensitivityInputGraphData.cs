using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class SensitivityInputGraphData
    {
        public int? Id { get; set; }
        public int ProjectId { get; set; }
        public string LineItem { get; set; }
        public string LineItemIntervals { get; set; }
        public string LineItemNPV { get; set; }
    }
}
