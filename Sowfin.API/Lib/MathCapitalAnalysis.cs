using Newtonsoft.Json;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static Sowfin.Model.Entities.CapitalAnalysis;
using static Sowfin.API.Lib.UnitConversion;
using Sowfin.API.ViewModels.CapitalStructure;
using static Sowfin.API.ViewModels.CapitalStructure.MathClass;

namespace Sowfin.API.Lib
{
    public class MathCapitalAnalysis
    {
        public static Dictionary<string, object> ScenarioAnalysis(CapitalAnalysisViewModel scenarioAnalysis, BaseSummaryOutput summaryOutput, double marginalTax)
        {
            double equityValue = summaryOutput.marketValueEquity; // from cap struct   Market Value of Common Equity (E) ($M)
            double preffEquityValue = summaryOutput.marketValuePreferredEquity;  /// Market Value of Preferred Equity(P)($M)
            double marketValueDebt = summaryOutput.marketValueDebt; //// Market Value of Debt (D) ($M)
            double cashEquivalent = summaryOutput.cashEquivalent;////Coming from cap strcut
            double excessCash = summaryOutput.excessCash;////Coming from cap strcut
            double unleveredCostCap = (summaryOutput.unleveredCostCapital / 100); ///Coming from cap strcut
            double costOfPreffEquity = (summaryOutput.costPreferredEquity / 100); //Coming from cap strcut

            ///summary out display
            double equityValueOut = 0;
            double preferredEquityValueOut = 0;
            double debtValue = 0;
            double changeDebt = 0;
            double netDebtValue = 0;
            double costOfEquity = 0;
            double costOfPreffEquityOut = 0;
            double costOfDebt = 0;
            double unleveredCostOut = 0;
            double weighedAverage = 0;
            double debtToEquity = 0;
            double debtToValue = 0;

            if (scenarioAnalysis.LeveragePolicy == "Constant Debt-to-Equity Ratio  (Target Leverage Ratio)")
            {
                ConstantDebtEquity constantDebtEquity = JsonConvert.DeserializeObject<ConstantDebtEquity>(scenarioAnalysis.AnalysisObject);

                equityValueOut = equityValue;
                preferredEquityValueOut = preffEquityValue;

                debtValue = equityValue * (constantDebtEquity.TargetDebtEquity / 100);
                changeDebt = debtValue - marketValueDebt;
                cashEquivalent = changeDebt + cashEquivalent;
                excessCash = changeDebt + excessCash;
                netDebtValue = debtValue - excessCash;

                costOfEquity = unleveredCostCap + ((netDebtValue / equityValue) * (unleveredCostCap - (constantDebtEquity.CostOfDebt / 100))) +
                    ((preffEquityValue / equityValue) * (unleveredCostCap - costOfPreffEquity));
                double sum = equityValue + preffEquityValue + netDebtValue;
                costOfPreffEquityOut = costOfPreffEquity;
                costOfDebt = constantDebtEquity.CostOfDebt / 100;
                unleveredCostOut = unleveredCostCap;
                weighedAverage = (costOfEquity * (equityValue / (sum))) + (costOfPreffEquity * (preffEquityValue / (sum))) +
                    ((costOfDebt * (netDebtValue / (sum))) * (1 - (marginalTax / 100)));
                debtToEquity = debtValue / equityValue;
                debtToValue = debtValue / (equityValue + preffEquityValue + debtValue);
            }
            else if (scenarioAnalysis.LeveragePolicy == "Annually Adjust Debt to Bring it Back in Line with Target Leverage Ratio")
            {
                ConstantDebtEquity constantDebtEquity = JsonConvert.DeserializeObject<ConstantDebtEquity>(scenarioAnalysis.AnalysisObject);
                equityValueOut = equityValue;
                preferredEquityValueOut = preffEquityValue;

                debtValue = equityValue * (constantDebtEquity.TargetDebtEquity / 100);
                changeDebt = debtValue - marketValueDebt;
                cashEquivalent = changeDebt + cashEquivalent;
                excessCash = changeDebt + excessCash;
                netDebtValue = debtValue - excessCash;

                costOfEquity = unleveredCostCap + ((netDebtValue / equityValue) * (unleveredCostCap - (constantDebtEquity.CostOfDebt / 100))) +
                    ((preffEquityValue / equityValue) * (unleveredCostCap - costOfPreffEquity));
                costOfPreffEquityOut = costOfPreffEquity;// converting at the top
                costOfDebt = constantDebtEquity.CostOfDebt / 100;
                unleveredCostOut = unleveredCostCap;
                weighedAverage = unleveredCostCap - (netDebtValue / (equityValue + preffEquityValue + netDebtValue)) * (marginalTax / 100) * costOfDebt *
                    ((1 + unleveredCostCap) / (1 + costOfDebt));
                debtToEquity = debtValue / equityValue;
                debtToValue = debtValue / (equityValue + preffEquityValue + debtValue);
            }
            else if (scenarioAnalysis.LeveragePolicy == "Constant Permanent Debt")
            {
                ConstPermanentDebt constPermanentDebt = JsonConvert.DeserializeObject<ConstPermanentDebt>(scenarioAnalysis.AnalysisObject);
                equityValueOut = equityValue;
                preferredEquityValueOut = preffEquityValue;
                unleveredCostOut = unleveredCostCap;

                debtValue = UnitConversion.ConvertUnits(new[] { constPermanentDebt.ValueOfPermDebt }, new[] { scenarioAnalysis.PermanentDebtUnit }, 0)[0];
                changeDebt = debtValue - marketValueDebt;

                cashEquivalent = changeDebt + cashEquivalent;
                excessCash = changeDebt + excessCash;

                netDebtValue = debtValue - excessCash;
                costOfDebt = (constPermanentDebt.CostOfDebt / 100);
                debtToEquity = debtValue / equityValue;
                debtToValue = debtValue / (equityValue + preffEquityValue + debtValue);

            }
            else if (scenarioAnalysis.LeveragePolicy == "Constant Interest Coverage Ratio (% of Free Cash Flow)")
            {
                ConstInterestCovRatio constPermanentDebt = JsonConvert.DeserializeObject<ConstInterestCovRatio>(scenarioAnalysis.AnalysisObject);
                equityValueOut = equityValue;
                preferredEquityValueOut = preffEquityValue;
                unleveredCostOut = unleveredCostCap;
                var freeCashFlow = UnitConversion.ConvertUnits(new[] { constPermanentDebt.FreeCashFlow }, new[] { scenarioAnalysis.FreeCashFlowUnit }, 0)[0];
                debtValue = ((constPermanentDebt.InterestCovRatio / 100) * freeCashFlow) / (constPermanentDebt.CostOfDebt / 100);
                changeDebt = debtValue - marketValueDebt;

                cashEquivalent = changeDebt + cashEquivalent;
                excessCash = changeDebt + excessCash;

                netDebtValue = debtValue - excessCash;
                costOfDebt = constPermanentDebt.CostOfDebt / 100;

                debtToEquity = debtValue / equityValue;
                debtToValue = debtValue / (equityValue + preffEquityValue + debtValue);

            }

            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("equityValue", AssignUnitCurrency(equityValueOut));
            result.Add("preferredEquityValue", AssignUnitCurrency(preferredEquityValueOut));
            result.Add("debtValue", AssignUnitCurrency(debtValue));
            result.Add("changeInDebt", AssignUnitCurrency(changeDebt));
            result.Add("netDebtValue", AssignUnitCurrency(netDebtValue));
            result.Add("cashEquivalent", AssignUnitCurrency(cashEquivalent));
            result.Add("excessCash", AssignUnitCurrency(excessCash));
            if (scenarioAnalysis.LeveragePolicy != "Constant Permanent Debt" || scenarioAnalysis.LeveragePolicy != "Constant Interest Coverage Ratio (% of Free Cash Flow)")
            {
                result.Add("costOfEquity", (costOfEquity * 100));
                result.Add("costOfPreferredEquity", (costOfPreffEquityOut * 100));
                result.Add("weightedAverage", (weighedAverage * 100));
            }
            result.Add("costOfDebt", (costOfDebt * 100));
            result.Add("unleveredCostOfCapital", (unleveredCostOut * 100));
            result.Add("debtToEquity", (debtToEquity * 100));
            result.Add("debtToValue", (debtToValue * 100));


            return result;

        }



        
    }
}
