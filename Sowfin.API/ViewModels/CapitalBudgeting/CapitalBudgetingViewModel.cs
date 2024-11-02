using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels
{
    public partial class CapitalBudgetingViewModel
    {
        public long Id { get; set; }
        public int StartingYear { get; set; }
        public int NoOfYears { get; set; }
        public double DiscountRate { get; set; }
        public double MarginalTaxRate { get; set; }
        public string TableData { get; set; }
        public long UserId { get; set; }
        public int ApprovalFlag { get; set; }
        public string SummaryOutput { get; set; }
        public string RawTableData { get; set; }
        public long ProjectId { get; set; }
        public int RevenueCount { get; set; }
        public int CapexCount { get; set; }
        public int FixedCostCount { get; set; }

        public CapitalBudgetingViewModel(int startingYear, int noOfYears, double discountRate, double marginalTaxRate, string tableData,
            long userId, int approvalFlag, string summaryOutput, string rawTableData, long projectId, int revenueCount,
            int capexCount, int fixedCostCount)
        {
            StartingYear = startingYear;
            NoOfYears = noOfYears;
            DiscountRate = discountRate;
            MarginalTaxRate = marginalTaxRate;
            TableData = tableData;
            UserId = userId;
            ApprovalFlag = approvalFlag;
            SummaryOutput = summaryOutput;
            RawTableData = rawTableData;
            ProjectId = projectId;
            RevenueCount = revenueCount;
            CapexCount = capexCount;
            FixedCostCount = fixedCostCount;


        }

        public CapitalBudgetingViewModel(long id, int startingYear, int noOfYears, double discountRate, double marginalTaxRate,
            string tableData, long userId, int approvalFlag, string summaryOutput, string rawTableData, long projectId,
            int revenueCount, int capexCount, int fixedCostCount)
        {
            Id = id;
            StartingYear = startingYear;
            NoOfYears = noOfYears;
            DiscountRate = discountRate;
            MarginalTaxRate = marginalTaxRate;
            TableData = tableData;
            UserId = userId;
            ApprovalFlag = approvalFlag;
            SummaryOutput = summaryOutput;
            RawTableData = rawTableData;
            ProjectId = projectId;
            RevenueCount = revenueCount;
            CapexCount = capexCount;
            FixedCostCount = fixedCostCount;

        }
    }
}
