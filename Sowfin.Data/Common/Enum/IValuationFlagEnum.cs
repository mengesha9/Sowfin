using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Sowfin.Data.Common.Enum
{
    public enum IValuationFlagEnum
    {
        [Description("Raw Historical Data")]
        RawHistorical = 1,
        [Description("Data Processing")]
        DataProcessing = 2,
        [Description("PayoutPolicy")]
        PayoutPolicy = 3,
        [Description("Interest")]
        Interest = 4,
        [Description("Tax")]
        Tax = 5,
        [Description("Cost of Capital")]
        costofCapital = 6,
        [Description("Integrated Financial Statement")]
        IntegratedStatement = 7,
        [Description("Hist. Analysis and Forcast Ratios")]
        ForcastRatio = 8,
        [Description("Reorganized Financial Statement")]
        ReorganizedStatement = 9,
        [Description("ROIC")]
        ROIC = 10,
        [Description("Non-Operating Assets and Non-Equity Claims")]
        NonOperatingAssets = 11,
        [Description("Valuation Summary")]
        ValuationSummary = 12,
        [Description("Sensitivity Analysis & Scenario Analysis")]
        Sensitivity = 13
        //Sensitivity Analysis & Scenario Analysis
    }
}
