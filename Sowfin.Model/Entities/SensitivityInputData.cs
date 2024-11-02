using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public class SensitivityInputData
    {
        // public int? Id { get; set; }
        // public List<SensitivityInputs> SensitivityInputValues { get; set; }

        public int? Id { get; set; }
        public int ProjectId { get; set; }
        public double ActualNPV { get; set; }
        public double NewNPV { get; set; }
        public double NewValue { get; set; }
        public double ActualValue { get; set; }
        public double IntervalDifference { get; set; }
        public string LineItem { get; set; }
        public string LineItemUnit { get; set; }
        public string NPVUnit { get; set; }
        public string LineItemIntervals { get; set; }
        public string LineItemNPV { get; set; }
    }

    //public class SensitivityInputs
    //{
    //    public int? Id { get; set; }
    //    public int ProjectId { get; set; }
    //    public double ActualNPV { get; set; }
    //    public double NewNPV { get; set; }
    //    public double NewValue { get; set; }
    //    public double ActualValue { get; set; }
    //    public double IntervalDifference { get; set; }
    //    public string LineItem { get; set; }
    //    public int LineItemUnit { get; set; }
    //    public int NPVUnit { get; set; }
    //    public string LineItemIntervals { get; set; }
    //    public string LineItemNPV { get; set; }
    //}

}
