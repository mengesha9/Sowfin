using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.ViewModels.InternalValuation
{
    public class PayoutPolicy_IValuationViewModel
    {
        public long Id { get; set; }
        public string WeightedAvgShares_Basic { get; set; }
        public string WeightedAvgShares_Diluted { get; set; }
        public string Annual_DPS { get; set; }
        public string TotalAnnualDividendPayout { get; set; }
        public string OneTimeDividendPayout { get; set; }
        public string StockPayBackAmount { get; set; }
        public string SharePurchased { get; set; }
        public string TableData { get; set; }
        public long? InitialSetupId { get; set; }
        public long UserId { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public int Yearcount { get; set; }
        public string Unit_WASOBasic { get; set; }
        public string Unit_WASODiluted { get; set; }
        public string Unit_ShareRepurchased { get; set; }
        public string Unit_TotalOngoingDividend { get; set; }
        public string Unit_OneTimeDividend { get; set; }
        public string Unit_StockBuyback { get; set; }


    }
}
