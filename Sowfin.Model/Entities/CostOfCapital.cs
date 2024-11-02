namespace Sowfin.Model.Entities
{
    public partial class CostOfCapital
    {
        public long Id { get; set; }
        public double MarketValueStock { get; set; }
        public double TotalValueStock { get; set; }
        public double MarketValueDebt { get; set; }
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
    }
}
