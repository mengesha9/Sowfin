using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels
{
    public class CapexBody
    {
        
        public long id { get; set; }
        public int duration { get; set; }
        public int add_to_year  { get; set; }
        public long projectId { get; set; }
        public double[] depreciationArray   { get; set; }
        public double[] customPercent { get; set; }
        public double[][] customPercenList { get; set; }
        public double[] add_to_year_array { get; set; }
    }
}