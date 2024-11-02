using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.CapitalBudgeting
{
    public class EvalSummaryOutput
    {
 
        public object[] Cogs { get; set; } = Array.Empty<object>();
        public object[] GrossMargins { get; set; } = Array.Empty<object>();
        public object[] SellingGenerals { get; set; } = Array.Empty<object>();
        public object[] RAndDs { get; set; } = Array.Empty<object>();
        public object[] OperatingIncomes { get; set; } = Array.Empty<object>();
        public object[] IncomeTaxs { get; set; } = Array.Empty<object>();
        public object[] UnleveredNetIcomes { get; set; } = Array.Empty<object>();
        public object[] Nwcs { get; set; } = Array.Empty<object>();
        public object[] IncreaseNwcs { get; set; } = Array.Empty<object>();
        public object[] FreeCashFlows { get; set; } = Array.Empty<object>();
        public object[] LeveredValues { get; set; } = Array.Empty<object>();
        public object[] DiscountFactors { get; set; } = Array.Empty<object>();
        public object[] DiscountCashFlow { get; set; } = Array.Empty<object>();
        public object[] Npv { get; set; } = Array.Empty<object>();
    }
}


