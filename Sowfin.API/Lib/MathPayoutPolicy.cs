using Sowfin.API.ViewModels.CapitalStructure;
using Sowfin.API.ViewModels.PayoutPolicy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sowfin.API.Lib
{
    public class MathPayoutPolicy
    {
        public static Dictionary<string, object> SummaryOutput(PayoutPolicyViewModel payoutPolicy, SummaryOutput summaryOuput)
        {
            double cashEquivalent = payoutPolicy.CashEquivalent - payoutPolicy.TotalOngoingPayout - payoutPolicy.OneTimePayout -
                 payoutPolicy.StockBuyBack;
            double excessCash = cashEquivalent - payoutPolicy.CashNeededWorkingCap;

            double totalPayoutQuarter = payoutPolicy.TotalOngoingPayout;
            double dpsQuarter = totalPayoutQuarter / payoutPolicy.NumberSharesBasic;
            double dividendPayoutQuater = totalPayoutQuarter / payoutPolicy.NetEarings;
            double dividendYieldQuater = dpsQuarter / payoutPolicy.currentSharePrice;

            double totalPayoutOne = payoutPolicy.OneTimePayout;
            double dpsOne = totalPayoutOne / payoutPolicy.NumberSharesBasic;
            double dividendPayoutOne = totalPayoutOne / payoutPolicy.NetEarings;
            double dividendYieldOne = dpsOne / payoutPolicy.currentSharePrice;

            double totalPayoutStock = payoutPolicy.StockBuyBack;
            double sharesRepurchased = totalPayoutStock / payoutPolicy.currentSharePrice;


            double netSales = payoutPolicy.NetSales;
            double ebitda = payoutPolicy.Ebitda;
            double deprectionAmortization = payoutPolicy.DepreciationAmortization;
            double interestIncome = cashEquivalent * (payoutPolicy.InterestIncomeRate / 100);
            double ebit = ebitda - deprectionAmortization + interestIncome;
            double interestExpense = payoutPolicy.TotalDebt * (payoutPolicy.CostOfDebt / 100);
            double ebt = ebit - interestExpense;
            double taxes = ebt * (payoutPolicy.MarginalTax / 100);
            double netEarnings = ebt - taxes;


            double sharesOutstandingBasic = payoutPolicy.NumberSharesBasic - sharesRepurchased;
            double sharesOutstandingDiluted = payoutPolicy.NumberSharesDiluted - sharesRepurchased;
            double netEarningsBasic = netEarnings / sharesOutstandingBasic;
            double netEarningsDiluted = netEarnings / sharesOutstandingDiluted;

            double finddiff = cashEquivalent - payoutPolicy.CashEquivalent;
            double totalCurrentAssets = payoutPolicy.TotalCurrentAssets + finddiff;
            double totalAssets = payoutPolicy.TotalAssests + finddiff;
            double totalDebt = payoutPolicy.TotalDebt;
            double shareHolderEquity = payoutPolicy.ShareholderEquity + finddiff;
            double cashFlowOperation = payoutPolicy.CashFlowOperation + (netEarnings - payoutPolicy.NetEarings);

            var output = UnitConversion.ConvertOutputUnits(out _, summaryOuput.marketValueDebt, summaryOuput.marketValueEquity);

            double debtMarketEquity = output[0] / output[1];
            double debtBookEquity = payoutPolicy.TotalDebt / shareHolderEquity;
            double ebitInterest = ebit / interestExpense;
            double ebitaInterest = ebitda / interestExpense;
            double cashFlowOperationDebt = cashFlowOperation / payoutPolicy.TotalDebt;
            double totalDebtEbita = payoutPolicy.TotalDebt / ebitda;

            return new Dictionary<string, object> {
                { "marketValueEquity", summaryOuput.marketValueEquity},
                { "marketValuePreferredEquity", summaryOuput.marketValuePreferredEquity},
                { "marketValueDebt", summaryOuput.marketValueDebt},
                { "marketValueNetDebt", summaryOuput.marketValueNetDebt},
                { "costEquity", (summaryOuput.costEquity) },
                { "costPreferredEquity", (summaryOuput.costPreferredEquity) },
                { "costDebt", (summaryOuput.costDebt) },
                { "unleveredCostCapital", (summaryOuput.unleveredCostCapital) },
                { "weighedAverageCapital", (summaryOuput.weighedAverageCapital) },
                { "cashEquivalent",  UnitConversion.AssignUnitCurrency(cashEquivalent) },
                { "excessCash", UnitConversion.AssignUnitCurrency(excessCash) },
                { "debtEquityRatio", (summaryOuput.debtEquityRatio) },
                { "debtValueRatio", (summaryOuput.debtValueRatio) },

                {"totalPayout", UnitConversion.AssignUnitCurrency(totalPayoutQuarter) },
                {"dpsQuarter", UnitConversion.AssignUnitCurrency(dpsQuarter)},
                {"dividendPayoutQuater",(dividendPayoutQuater*100) },
                {"dividendYieldQuater",(dividendYieldQuater*100) },

                {"totalPayoutOne",UnitConversion.AssignUnitCurrency(totalPayoutOne) },
                {"dpsOne",UnitConversion.AssignUnitCurrency(dpsOne) },
                {"dividendPayoutOne",(dividendPayoutOne*100) },
                {"dividendYieldOne",(dividendYieldOne*100) },

                {"totalPayoutStock", UnitConversion.AssignUnitCurrency(totalPayoutStock) },
                {"sharesRepurchased", UnitConversion.AssignUnit(sharesRepurchased)},

                {"sharesOutstandingBasic",UnitConversion.AssignUnit(sharesOutstandingBasic) },
                {"sharesOutstandingDiluted",UnitConversion.AssignUnit(sharesOutstandingDiluted) },
                {"netEarningsBasic",UnitConversion.AssignUnitCurrency(netEarningsBasic) },
                {"netEarningsDiluted",UnitConversion.AssignUnitCurrency(netEarningsDiluted) },


                {"netSales" , UnitConversion.AssignUnitCurrency(netSales) },
                { "ebitda" , UnitConversion.AssignUnitCurrency(ebitda) },
                { "deprectionAmortization" ,UnitConversion.AssignUnitCurrency(deprectionAmortization) },
                { "interestIncome" ,UnitConversion.AssignUnitCurrency(interestIncome) },
                { "ebit" ,UnitConversion.AssignUnitCurrency(ebit) },
                { "interestExpense" ,UnitConversion.AssignUnitCurrency(interestExpense) },
                { "ebt" , UnitConversion.AssignUnitCurrency(ebt) },
                { "taxes" ,UnitConversion.AssignUnitCurrency(taxes) },
                { "netEarnings" ,UnitConversion.AssignUnitCurrency(netEarnings)},

                {"cashEquivalents", UnitConversion.AssignUnitCurrency(cashEquivalent) },
                {"totalCurrentAssets" , UnitConversion.AssignUnitCurrency(totalCurrentAssets) },
                {"totalAssets" ,  UnitConversion.AssignUnitCurrency(totalAssets) },
                {"totalDebt",  UnitConversion.AssignUnitCurrency(totalDebt) },
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
