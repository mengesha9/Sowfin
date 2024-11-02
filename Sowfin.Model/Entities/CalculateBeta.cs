using System;

namespace Sowfin.Model.Entities
{
    public class CalculateBeta
    {
        public long Id { get; set; }
        public long? CostOfCapitals_Id { get; set; }
        public long? MasterId { get; set; }
        public long? Frequency_Id { get; set; }
        public long? TargetMarketIndex_Id { get; set; }
        public long? TargetRiskFreeRate_Id { get; set; }
        public long? DataSource_Id { get; set; }
        public DateTime? Duration_FromDate { get; set; }
        public DateTime? Duration_toDate { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public bool Active { get; set; }
        public double? BetaValue { get; set; }
        public string FileData { get; set; }
    }
}
