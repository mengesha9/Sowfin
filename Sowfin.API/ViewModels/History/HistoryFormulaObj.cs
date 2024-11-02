using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.History
{
    public class HistoryFormulaObj
    {
        public List<List<string>> payoutTableAddr { get; set; }
        public List<List<string>> shareTableAddr  { get; set; }
        public List<List<string>> currentCapitalTableAddr { get; set; }
        public List<List<string>> currentCostOfCapTableAddr  { get; set; }
        public List<List<string>> otherInputTableAddr  {get;set;}
        public List<List<string>> financialTableAddr1 {get;set;}
        public List<List<string>> financialTableAddr2 {get;set;}
        public List<List<string>> financialTableAddr3 {get;set;}
    }
}
