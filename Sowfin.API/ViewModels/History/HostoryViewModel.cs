using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.History
{
    public class HostoryViewModel
    {

        public long Id { get; set; }
        public long UserId { get; set; }
        public string PayoutTable { get; set; }
        public string ShareTable { get; set; }
        public string CurrentCapitalTable { get; set; }
        public string CurrentCostOfCapTable { get; set; }
        public string OtherInputTable { get; set; }
        public string FinancialTable { get; set; }
        public string SummaryOutput { get; set; }
        public int ApprovalFlag { get; set; }
        public int StartYear { get; set; }
        public int EndYear { get; set; }
        public int SummaryFlag { get; set; }
        public HostoryViewModel()
        {
        }

        public HostoryViewModel(long userId, string payoutTable, string shareTable, string currentCapitalTable,
            string currentCostOfCapTable, string otherInputTable, string financialTable, string summaryOutput, int startYear,
            int endYear, int summaryFlag)
        {
            UserId = userId;
            PayoutTable = payoutTable;
            ShareTable = shareTable;
            CurrentCapitalTable = currentCapitalTable;
            CurrentCostOfCapTable = currentCostOfCapTable;
            OtherInputTable = otherInputTable;
            FinancialTable = financialTable;
            SummaryOutput = summaryOutput;
            StartYear = startYear;
            EndYear = endYear;
            SummaryFlag = summaryFlag;
        }

        public HostoryViewModel(long id, long userId, string payoutTable, string shareTable, string currentCapitalTable,
            string currentCostOfCapTable, string otherInputTable, string financialTable, string summaryOutput, int startYear,
            int endYear, int summaryFlag)
        {
            Id = id;
            UserId = userId;
            PayoutTable = payoutTable;
            ShareTable = shareTable;
            CurrentCapitalTable = currentCapitalTable;
            CurrentCostOfCapTable = currentCostOfCapTable;
            OtherInputTable = otherInputTable;
            FinancialTable = financialTable;
            SummaryOutput = summaryOutput;
            StartYear = startYear;
            EndYear = endYear;
            SummaryFlag = summaryFlag;
        }



    }
}
