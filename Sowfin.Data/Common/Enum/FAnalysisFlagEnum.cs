using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum FAnalysisFlagEnum
    {
        [Description("Raw Historical Data")]
        RawHistorical = 1,
        [Description("Data Processing")]
        DataProcessing = 2,
        [Description("Market Data")]
        MarketData = 3,
        [Description("Integrated Financial Statement")]
        IntegratedStatement = 4,
        [Description("Financial Statement Analysis")]
        FinancialAnalysis = 5
    }
}
