using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalBudgeting
{
    public class AggregateList
    {
        public List<double> Volume { get; set; }
        public List<double> UnitPrice  { get; set; }
        public List<double> UnitCost  { get; set; }
        public List<double> Fixed  { get; set; }
        public List<double> NWC   { get; set; }
        public List<double> Capex  { get; set; }
        public List<double> TotalDepreciation  { get; set; }
    }
}
