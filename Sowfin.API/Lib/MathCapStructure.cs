using Sowfin.API.ViewModels;
using System;
using System.Collections.Generic;
using static Sowfin.API.Lib.UnitConversion;

namespace Sowfin.API.Lib
{
    public class MathCapStructure
    {
        public static Dictionary<string, object> SummaryOutput(CapitalStructureViewModel capitalStructure)
        {
            double cash_equivalent = capitalStructure.CashEquivalent;  //8000.00;   //D21  //M15  //V15  //AF15
            double cash_Needed_wc = capitalStructure.CashNeededCapital; //3000.00;   //D22   //M16  //V16   //AF16
            double interest_coverage = (capitalStructure.InterestCoverage / 100); //0;  //D23  //M17  //V17  //AF17
            double marginal_tax_rate = (capitalStructure.MarginalTaxRate / 100); //(26.70 / 100);  //D24  //M18  //V18   //AF18
            double free_cash_flow = capitalStructure.FreeCashFlow;  //0;// //D19  ///For 4 method  //AF19


            ///Equity
            double currentSharePrice = (capitalStructure.equity!=null ? capitalStructure.equity.CurrentSharePrice :0); //31.30;
            double sharesOutstandingBasic = (capitalStructure.equity != null ? capitalStructure.equity.NumberShareBasic :0); //4951.00;// 
            double sharesOutstandingDiluted = (capitalStructure.equity != null ? capitalStructure.equity.NumberShareOutstanding :0); //5500.00;//
            double costOfEquity = (capitalStructure.equity != null ? (capitalStructure.equity.CostOfEquity / 100) :0); // (9.78 / 100); //

            double costOfDebt = (capitalStructure.debt!= null ? (capitalStructure.debt.CostOfDebt / 100) :0); // (2.27 / 100);  //D12 Method 1  //M12  Method 2  //V12  //AF12
            double marketValueOfID = (capitalStructure.debt != null ? (capitalStructure.debt.MarketValueDebt) :0);  //16194.00;//  //D11 Method 1  //M11  Method 2 //V11 

            double currentPreferredSharePrice = (capitalStructure.prefferedEquity!=null ? capitalStructure.prefferedEquity.PrefferedSharePrice :0); // 10.00;//;
            double numberOfpreferredSharesOutstanding = (capitalStructure.prefferedEquity != null ? capitalStructure.prefferedEquity.PrefferedShareOutstanding :0);// 220.00;// ;
            double preferredDividend = (capitalStructure.prefferedEquity != null ?  capitalStructure.prefferedEquity.PrefferedDividend :0);  //1.00;//
            double costofPreferredEquity = (capitalStructure.prefferedEquity != null ? (capitalStructure.prefferedEquity.CostPreffEquity / 100) :0);  //(10.00 / 100); //

            /////These formulas remain the same for all the methods  //// To be showed in summary output
            double marketValueEquity = currentSharePrice * sharesOutstandingBasic;
            double marketValuePreffEquity = currentPreferredSharePrice * numberOfpreferredSharesOutstanding;
            double marketValueDebt = marketValueOfID;
            double marketValueNetDebt = marketValueDebt - (cash_equivalent - cash_Needed_wc);
            double costEquity = costOfEquity;
            double costPreffEquity = costofPreferredEquity;
            double costDebt = costOfDebt;


            double sumOfInputs = marketValueEquity + marketValuePreffEquity + marketValueNetDebt;
            double unleveredCostCapital = ((costOfEquity * (marketValueEquity / (sumOfInputs))) + (costofPreferredEquity * (marketValuePreffEquity / (sumOfInputs))) + (costOfDebt * (marketValueNetDebt / (sumOfInputs))));
            double weighedAverageCapital = ((costOfEquity * (marketValueEquity / (sumOfInputs))) + (costofPreferredEquity * (marketValuePreffEquity / (sumOfInputs))) + (costOfDebt * (marketValueNetDebt / (sumOfInputs)) * (1 - marginal_tax_rate)));

            double cashEquivalent = cash_equivalent;
            double excessCash = cashEquivalent - cash_Needed_wc;
            double debtEquityRatio = marketValueDebt / marketValueEquity;
            double debtValueRatio = (marketValueDebt / (marketValueEquity + marketValuePreffEquity + marketValueDebt));/// this remains same for first three methods change for third method
                                                                                                                       ///Not clear
            double unleveredEnterpriseValue = 0;
            double leverateEnterpriseValue = 0;
            double equityValue = 0;
            double interestTaxShield = 0;

            if(capitalStructure.NewLeveragePolicy != null )
            if (capitalStructure.NewLeveragePolicy == "Constant Debt-to-Equity Ratio  (Target Leverage Ratio)")
            {         
                    /////These formulae doubt
                unleveredEnterpriseValue = excessCash; //D41 //M41  //V41 //AF41 doubt
                leverateEnterpriseValue = excessCash;  //D42 //M42  //doubt for 1st and 2nd method
                equityValue = leverateEnterpriseValue - marketValueOfID; //D43 //Method 1
                interestTaxShield = leverateEnterpriseValue - unleveredEnterpriseValue; //D44  //M44 
            }
            else if (capitalStructure.NewLeveragePolicy == "Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")
            {
                weighedAverageCapital = unleveredCostCapital - ((marketValueNetDebt / (marketValueEquity + marketValuePreffEquity + marketValueNetDebt)) * marginal_tax_rate * costDebt * ((1 + unleveredCostCapital) / (1 + costDebt)));
                /////These formulae doubt
                unleveredEnterpriseValue = excessCash; //D41 //M41  //V41 //AF41 doubt
                leverateEnterpriseValue = excessCash;  //D42 //M42  //doubt for 1st and 2nd method
                equityValue = leverateEnterpriseValue - marketValueDebt; //M43  //Method 2
                interestTaxShield = leverateEnterpriseValue - unleveredEnterpriseValue; //D44  //M44
            }
            else if (capitalStructure.NewLeveragePolicy == "Constant Permanent Debt")
            {
                weighedAverageCapital = 0;/// NA in excel
                /////These formulae doubt
                unleveredEnterpriseValue = excessCash; //D41 //M41  //V41 //AF41 doubt
                leverateEnterpriseValue = (unleveredEnterpriseValue + (marginal_tax_rate * free_cash_flow));  //V41 +(V18 *V24) //V42
                equityValue = leverateEnterpriseValue - marketValueDebt;  //V43 // AF43
                interestTaxShield = (marginal_tax_rate * free_cash_flow); //V44

            }
            else if (capitalStructure.NewLeveragePolicy == "Constant Interest Coverage Ratio (% of Free Cash Flow)")
            {
                weighedAverageCapital = 0;/// NA in excel
                debtEquityRatio = marketValueNetDebt / marketValueEquity;
                debtValueRatio = (marketValueNetDebt / (marketValueEquity + marketValuePreffEquity + marketValueDebt + marketValueNetDebt));
                /////These formulae doubt
                unleveredEnterpriseValue = excessCash; //D41 //M41  //V41 //AF41 doubt
                leverateEnterpriseValue = (unleveredEnterpriseValue + (marginal_tax_rate * interest_coverage * unleveredEnterpriseValue)); //AF42
                equityValue = leverateEnterpriseValue - marketValueDebt;  //V43 // AF43
                interestTaxShield = marginal_tax_rate * interest_coverage * unleveredEnterpriseValue; //AF44
            }


            double stockPrice = equityValue / sharesOutstandingBasic; //D45  //M45  //V45 //AF45
            Dictionary<string, object> result = new Dictionary<string, object>
            {   { "marketValueEquity", AssignUnitCurrency(marketValueEquity)},
                { "marketValuePreferredEquity",AssignUnitCurrency(marketValuePreffEquity)},
                { "marketValueDebt", AssignUnitCurrency(marketValueDebt)},
                { "marketValueNetDebt", AssignUnitCurrency(marketValueNetDebt)},
                { "costEquity", (costEquity*100) },
                { "costPreferredEquity", (costPreffEquity*100) },
                { "costDebt", (costDebt*100) },
                { "unleveredCostCapital", (unleveredCostCapital*100) },
                { "weighedAverageCapital", (weighedAverageCapital*100) },
                { "cashEquivalent", AssignUnitCurrency(cashEquivalent) },
                { "excessCash", AssignUnitCurrency(excessCash) },
                { "debtEquityRatio", (debtEquityRatio*100) },
                { "debtValueRatio", (debtValueRatio*100) },
                { "unleveredEnterpriseValue", unleveredEnterpriseValue },
                { "leverateEnterpriseValue", leverateEnterpriseValue },
                { "equityValue",equityValue },
                { "interestTaxShield", interestTaxShield },
                { "stockPrice", AssignUnitCurrency(stockPrice) }
            };
            return result;
        }


    }
}
