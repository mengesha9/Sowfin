using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels
{
    public class CostOfCapitalViewModel
    {
        public long Id { get; set; }
        public double MarketValueStock { get; set; }
        public double TotalValueStock { get; set; }
        public double MarketValueDebt { get; set; }
        public string FileData { get; set; }

        public double RiskFreeRate { get; set; }
        public double HistoricMarket { get; set; }
        public double HistoricRiskReturn { get; set; }
        public double SmallStock { get; set; }
        public double RawBeta { get; set; }

        public int ApprovalFlag { get; set; }

        public double PreferredDividend { get; set; }
        public double PreferredShare { get; set; }

        public double TaxRate { get; set; }   //%
        public double ProjectRisk { get; set; } //%
        public string Method { get; set; }
        public int MethodType { get; set; }
        public long UserId { get; set; }
        public string SummaryOutput { get; set; }

        public string MarketValueUnit { get; set; }
        public string TotalValueUnit { get; set; }
        public string MarketDebtUnit { get; set; }
        public string PreferredDividendUnit { get; set; }
        public string PreferredShareUnit { get; set; }
        public int SummaryFlag { get; set; }

        public string CompanyName { get; set; }

        // Calculate Beta Changes
        public long? CalculateBeta_Id { get; set; }
        public long? CostOfCapitals_Id { get; set; }
        public long? Frequency_Id { get; set; }
        public string Frequency_Value { get; set; }
        public long? TargetMarketIndex_Id { get; set; }
        public string TargetMarketIndex_Value { get; set; }
        public long? TargetRiskFreeRate_Id { get; set; }
        public string TargetRiskFreeRate_Value { get; set; }
        public long? DataSource_Id { get; set; }
        public string DataSource_Value { get; set; }
        public DateTime? Duration_FromDate { get; set; }
        public DateTime? Duration_toDate { get; set; }
        public DateTime? Beta_CreatedDate { get; set; }
        public DateTime? Beta_ModifiedDate { get; set; }
        public bool Beta_Active { get; set; }
        public double? BetaValue { get; set; }

        public long? MasterId { get; set; }


    }
    public class CalculateFileViewModel
    {
        public string FileData { get; set; }
        public long? CalculateBeta_Id { get; set; }


    }
}
