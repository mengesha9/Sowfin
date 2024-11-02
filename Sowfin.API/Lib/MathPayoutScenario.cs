using Newtonsoft.Json;
using Sowfin.API.ViewModels.CapitalStructure;
using Sowfin.API.ViewModels.PayoutPolicy;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.Lib
{
    public class MathPayoutScenario
    {
        public static Dictionary<string, object> SummaryOutput(PayoutScenario payoutScenario, PayoutPolicy payoutPolicy)
        {
            var summaryOutput = JsonConvert.DeserializeObject<SummaryOutput>(payoutPolicy.SummaryOutput);

            var output = UnitConversion.ConvertOutputUnits(out _, summaryOutput.marketValueEquity, summaryOutput.marketValueDebt,
                summaryOutput.marketValuePreferredEquity, summaryOutput.marketValuePreferredEquity);


            double debtValue = (payoutScenario.TargetDebtToEquity / 100) * output[0];
            double changeInDebt = debtValue - output[1];


            double cashEquivalent = payoutPolicy.CashEquivalent + changeInDebt - payoutScenario.TotalPayoutAnnual - payoutScenario.OneTimePayout - payoutScenario.StockBuyBack;
            double excessCash = cashEquivalent - payoutPolicy.CashNeededWorkingCap;

            double netDebtValue = debtValue - excessCash;

            double costOfEquity = (summaryOutput.unleveredCostCapital) + ((netDebtValue / output[0]) * ((summaryOutput.unleveredCostCapital) - (payoutScenario.CostOfDebt))) + ((output[3] / output[0]) * ((summaryOutput.unleveredCostCapital) - (payoutPolicy.CostOfPrefEquity)));
            double sum = output[0] + output[2] + netDebtValue;
            double weightedAverage = ((costOfEquity) * (output[0] / sum)) + ((payoutPolicy.CostOfPrefEquity) * (output[2] / sum)) + ((payoutScenario.CostOfDebt) * (netDebtValue / sum)) * (1 - (payoutPolicy.MarginalTax / 100));

            double debtEquity = debtValue / output[0];
            double debTotValue = debtValue / (output[0] + output[2] + debtValue);

            double totalPayoutQuarter = payoutScenario.TotalPayoutAnnual;
            double dpsQuarter = totalPayoutQuarter / payoutPolicy.NumberSharesBasic;
            double dividendPayoutQuarter = totalPayoutQuarter / payoutPolicy.NetEarings;
            double dividendYieldQuarter = dpsQuarter / payoutPolicy.currentSharePrice;

            double totalPayoutOne = payoutScenario.OneTimePayout;
            double dpsOne = totalPayoutOne / payoutPolicy.NumberSharesBasic;
            double dividendPayoutOne = totalPayoutOne / payoutPolicy.NetEarings;
            double dividendYieldOne = dpsOne / payoutPolicy.currentSharePrice;

            double totalPayoutStock = payoutScenario.StockBuyBack;
            double sharesRepurchased = totalPayoutStock / payoutPolicy.currentSharePrice;

            double netSales = payoutPolicy.NetSales;
            double ebitda = payoutPolicy.Ebitda;
            double deprectionAmortization = payoutPolicy.DepreciationAmortization;
            double interestIncome = cashEquivalent * (payoutPolicy.InterestIncomeRate / 100);
            double ebit = ebitda - deprectionAmortization + interestIncome;
            double interestExpense = debtValue * (payoutScenario.CostOfDebt / 100);
            double ebt = ebit + interestExpense;
            double taxes = ebt * (payoutPolicy.MarginalTax / 100);
            double netEarnings = ebt - taxes;


            double numberSharesBasic = payoutPolicy.NumberSharesBasic - sharesRepurchased;
            double numberSharesDiluted = payoutPolicy.NumberSharesDiluted - sharesRepurchased;
            double netEarningsBasic = netEarnings / numberSharesBasic;
            double netEarningsDiluted = netEarnings / numberSharesDiluted;

            double finddiff = cashEquivalent - payoutPolicy.CashEquivalent;
            double totalCurrentAssets = payoutPolicy.TotalCurrentAssets + finddiff;
            double totalAssets = payoutPolicy.TotalAssests + finddiff;
            double shareHolderEquity = payoutPolicy.ShareholderEquity + finddiff;
            double cashFlowOperation = payoutPolicy.CashFlowOperation + (netEarnings - payoutPolicy.NetEarings);

            double debtMarketEquity = debtValue / output[0];
            double debtBookEquity = debtValue / shareHolderEquity;
            double ebitInterest = ebit / interestExpense;
            double ebitaInterest = ebitda / interestExpense;
            double cashFlowOperationDebt = cashFlowOperation / debtValue;
            double totalDebtEbita = debtValue / ebitda;


            return new Dictionary<string, object> {
                { "marketValueEquity", summaryOutput.marketValueEquity},
                { "marketValuePreferredEquity", summaryOutput.marketValuePreferredEquity},
                { "marketValueDebt", UnitConversion.AssignUnitCurrency(debtValue)},
                { "changeInDebt" , UnitConversion.AssignUnitCurrency(changeInDebt)},
                { "marketValueNetDebt", UnitConversion.AssignUnitCurrency(netDebtValue)},
                { "costEquity", costOfEquity },
                { "costPreferredEquity", (summaryOutput.costPreferredEquity) },
                { "costDebt", payoutScenario.CostOfDebt },
                { "unleveredCostCapital", (summaryOutput.unleveredCostCapital) },
                { "weighedAverageCapital", weightedAverage },
                { "cashEquivalent", UnitConversion.AssignUnitCurrency(cashEquivalent) },
                { "excessCash", UnitConversion.AssignUnitCurrency(excessCash) },
                { "debtEquityRatio", (debtEquity*100) },
                { "debtValueRatio", (debTotValue*100) },

                {"totalPayout", UnitConversion.AssignUnitCurrency(totalPayoutQuarter) },
                {"dpsQuarter", UnitConversion.AssignUnitCurrency(dpsQuarter)},
                {"dividendPayoutQuater",(dividendPayoutQuarter*100) },
                {"dividendYieldQuater",(dividendYieldQuarter*100) },

                {"totalPayoutOne",UnitConversion.AssignUnitCurrency(totalPayoutOne) },
                {"dpsOne",UnitConversion.AssignUnitCurrency(dpsOne) },
                {"dividendPayoutOne",(dividendPayoutOne*100) },
                {"dividendYieldOne",(dividendYieldOne*100) },

                {"totalPayoutStock",UnitConversion.AssignUnitCurrency(totalPayoutStock) },
                {"sharesRepurchased", UnitConversion.AssignUnit(sharesRepurchased)},

                {"sharesOutstandingBasic",UnitConversion.AssignUnit(numberSharesBasic) },
                {"sharesOutstandingDiluted",UnitConversion.AssignUnit(numberSharesDiluted) },
                {"netEarningsBasic",UnitConversion.AssignUnitCurrency(netEarningsBasic) },
                {"netEarningsDiluted",UnitConversion.AssignUnitCurrency(netEarningsDiluted) },


                {"netSales" , UnitConversion.AssignUnitCurrency(netSales) },
                { "ebitda" ,UnitConversion.AssignUnitCurrency(ebitda) },
                { "deprectionAmortization" ,UnitConversion.AssignUnitCurrency(deprectionAmortization) },
                { "interestIncome" ,UnitConversion.AssignUnitCurrency(interestIncome) },
                { "ebit" ,UnitConversion.AssignUnitCurrency(ebit) },
                { "interestExpense" ,UnitConversion.AssignUnitCurrency(interestExpense) },
                { "ebt" , UnitConversion.AssignUnitCurrency(ebt) },
                { "taxes" ,UnitConversion.AssignUnitCurrency(taxes) },
                { "netEarnings" ,UnitConversion.AssignUnitCurrency(netEarnings)},

                {"cashEquivalents",UnitConversion.AssignUnitCurrency(cashEquivalent) },
                {"totalCurrentAssets" , UnitConversion.AssignUnitCurrency(totalCurrentAssets) },
                {"totalAssets" ,  UnitConversion.AssignUnitCurrency(totalAssets) },
                {"totalDebt",  UnitConversion.AssignUnitCurrency(debtValue) },
                {"shareHolderEquity",  UnitConversion.AssignUnitCurrency(shareHolderEquity) },
                {"cashFlowOperation", UnitConversion.AssignUnitCurrency(cashFlowOperation) },

                {"debtMarketEquity" ,(debtMarketEquity*100) },
                {"debtBookEquity",(debtBookEquity*100) },
                { "ebitInterest",ebitInterest },
                {"ebitaInterest",ebitaInterest },
                {"cashFlowOperationDebt" ,( cashFlowOperationDebt*100) },
                {"totalDebtEbita" ,totalDebtEbita},
            };
        }
    }
}
