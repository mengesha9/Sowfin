using Sowfin.API.ViewModels.History;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.Lib
{
    public class MathHistory
    {
        public static HistorySummary SummaryOuput(HistoryObject historyObject)
        {
            HistorySummary historySummary = new HistorySummary();
            for (int i = 2; i < historyObject.payoutTable[0].Length; i++)
            {
                var totalCashReturn = ParseDouble(historyObject.payoutTable[0][i]) + ParseDouble(historyObject.payoutTable[1][i]) + ParseDouble(historyObject.payoutTable[2][i]);
                var totalCashReturnFCFE = totalCashReturn / ParseDouble(historyObject.otherInputTable[1][i]);
                var totalPayoutQuarter = ParseDouble(historyObject.payoutTable[0][i]);
                var dpsQuarter = totalPayoutQuarter / ParseDouble(historyObject.shareTable[1][i]);
                var dividendPayoutQuarter = totalPayoutQuarter / ParseDouble(historyObject.financialTable[8][i]);
                var dividendYieldQuarter = dpsQuarter / ParseDouble(historyObject.shareTable[0][i]);

                totalCashReturnFCFE = Convert.ToDouble(totalCashReturnFCFE.ToString("0.##"));
                dpsQuarter = Convert.ToDouble(dpsQuarter.ToString("0.##"));
                dividendYieldQuarter = Convert.ToDouble(dividendYieldQuarter.ToString("0.##"));

                historySummary.TotalCashReturned.Add(UnitConversion.AssignUnitCurrency(totalCashReturn));
                historySummary.TotalCashReturnedFCFE.Add(totalCashReturnFCFE * 100);

                historySummary.TotalPayoutQuarter.Add(UnitConversion.AssignUnitCurrency(totalPayoutQuarter));
                historySummary.DPSQuarter.Add(UnitConversion.AssignUnitCurrency(dpsQuarter));
                historySummary.PayoutRatioQuarter.Add(dividendPayoutQuarter * 100);
                historySummary.DividendYieldQuarter.Add(Convert.ToDouble((dividendYieldQuarter * 100).ToString("0.##")) );


                var totalPayoutOne = ParseDouble(historyObject.payoutTable[1][i]);
                var dpsOne = totalPayoutOne / ParseDouble(historyObject.shareTable[1][i]);
                dpsOne = Convert.ToDouble(dpsOne.ToString("0.##"));


                var dividendPayoutOne = totalPayoutOne / ParseDouble(historyObject.financialTable[8][i]);
                var dividendYieldOne = dpsOne / ParseDouble(historyObject.shareTable[0][i]);
                dividendYieldOne = Convert.ToDouble(dividendYieldOne.ToString("0.##"));


                historySummary.TotalPayoutOne.Add(UnitConversion.AssignUnitCurrency(totalPayoutOne));
                historySummary.DPSOne.Add(UnitConversion.AssignUnitCurrency(dpsOne));
                historySummary.DividendPayoutRatioOne.Add(dividendPayoutOne * 100);
                historySummary.DividendYieldOne.Add(Convert.ToDouble((dividendYieldOne * 100).ToString("0.##")));


                var debtMarketEquity = ParseDouble(historyObject.currentCapitalTable[1][i]) / ParseDouble(historyObject.currentCapitalTable[0][i]);
                var debtBookEquity = ParseDouble(historyObject.financialTable[12][i]) / ParseDouble(historyObject.financialTable[13][i]);
                var ebitCoverage = ParseDouble(historyObject.financialTable[4][i]) / ParseDouble(historyObject.financialTable[5][i]);
                var ebitdaCoverage = ParseDouble(historyObject.financialTable[1][i]) / ParseDouble(historyObject.financialTable[5][i]);
                var cashFlowOperation = ParseDouble(historyObject.financialTable[14][i]) / ParseDouble(historyObject.financialTable[12][i]);
                var totalDebt = ParseDouble(historyObject.financialTable[12][i]) / ParseDouble(historyObject.financialTable[1][i]);

                cashFlowOperation = Convert.ToDouble(cashFlowOperation.ToString("0.##"));
                totalDebt = Convert.ToDouble(totalDebt.ToString("0.##"));

                historySummary.DebtMarketEquity.Add(debtMarketEquity * 100);
                historySummary.DebtBookEquity.Add(debtBookEquity * 100);
                historySummary.EBITInterestCoverage.Add(ebitCoverage);
                historySummary.EBITDAInterestCoverage.Add(ebitdaCoverage);
                historySummary.CashFlow.Add(cashFlowOperation * 100);
                historySummary.TotalDebt.Add(totalDebt);


                var excessCash = ParseDouble(historyObject.financialTable[9][i]) - ParseDouble(historyObject.otherInputTable[0][i]);
                historySummary.ExcessCash.Add(UnitConversion.AssignUnitCurrency(excessCash));

            }

            historySummary.TotalPayout = historyObject.payoutTable[2].Skip(2).Select(x => (object)UnitConversion.AssignUnitCurrency((double)x)).ToList();
            historySummary.SharesRepurchased = historyObject.payoutTable[3].Skip(2).Select(x => (object)UnitConversion.AssignUnit((double)x)).ToList();

            historySummary.CashEquivalent = historyObject.financialTable[9].Skip(2).Select(x => (object)UnitConversion.AssignUnitCurrency((double)x)).ToList();
            historySummary.CashNeededCapital = historyObject.otherInputTable[0].Skip(2).Select(x => (object)UnitConversion.AssignUnitCurrency((double)x)).ToList();




            return historySummary;
        }
        public static double ParseDouble(Object obj)
        {
            if (Convert.ToString(obj) == "" || obj == null)
            {
                obj = 0;
            }
            return Convert.ToDouble(obj);
        }
    }
}
