using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
    public partial class CapitalAnalysis
    {
        public long Id { get; set; }
        public string LeveragePolicy { get; set; }
        public string AnalysisObject { get; set; }
        public long UserId { get; set; }
        public string SummaryOutput { get; set; }
        public string PermanentDebtUnit { get; set; }
        public string FreeCashFlowUnit { get; set; }


        public class BaseSummaryOutput
        {
            public double marketValueEquity { get; set; }
            public double marketValuePreferredEquity { get; set; }
            public double marketValueDebt { get; set; }
            public double marketValueNetDebt { get; set; }
            public double costEquity { get; set; }
            public double costPreferredEquity { get; set; }
            public double costDebt { get; set; }
            public double unleveredCostCapital { get; set; }
            public double weighedAverageCapital { get; set; }
            public double cashEquivalent { get; set; }
            public double excessCash { get; set; }
            public double debtEquityRatio { get; set; }
            public double debtValueRatio { get; set; }
            public double unleveredEnterpriseValue { get; set; }
            public double leverateEnterpriseValue { get; set; }
            public double equityValue { get; set; }
            public double interestTaxShield { get; set; }
            public string stockPrice { get; set; }


            public BaseSummaryOutput()
            {

            }

            public BaseSummaryOutput(double marketValueEquity, double marketValuePreferredEquity, double marketValueDebt, double marketValueNetDebt, double costEquity, double costPreferredEquity, double costDebt, double unleveredCostCapital, double weighedAverageCapital, double cashEquivalent, double excessCash, double debtEquityRatio, double debtValueRatio, double unleveredEnterpriseValue, double leverateEnterpriseValue, double equityValue, double interestTaxShield, string stockPrice)
            {
                this.marketValueEquity = marketValueEquity;
                this.marketValuePreferredEquity = marketValuePreferredEquity;
                this.marketValueDebt = marketValueDebt;
                this.marketValueNetDebt = marketValueNetDebt;
                this.costEquity = costEquity;
                this.costPreferredEquity = costPreferredEquity;
                this.costDebt = costDebt;
                this.unleveredCostCapital = unleveredCostCapital;
                this.weighedAverageCapital = weighedAverageCapital;
                this.cashEquivalent = cashEquivalent;
                this.excessCash = excessCash;
                this.debtEquityRatio = debtEquityRatio;
                this.debtValueRatio = debtValueRatio;
                this.unleveredEnterpriseValue = unleveredEnterpriseValue;
                this.leverateEnterpriseValue = leverateEnterpriseValue;
                this.equityValue = equityValue;
                this.interestTaxShield = interestTaxShield;
                this.stockPrice = stockPrice;
            }
        }




    }
}
