using Newtonsoft.Json;
using Sowfin.API.ViewModels;
using System.Collections.Generic;

namespace Sowfin.API.Lib
{
    public class MathCostOfCapital
    {
        public static Dictionary<string, double> SummaryOutput(CostOfCapitalViewModel obj)
        {
            double marketValueCommon = obj.MarketValueStock;  //154867.23f;   //D9
            double totalValuePS = obj.TotalValueStock; // 2.2f;  //D10
            double marketValueOfND = obj.MarketValueDebt; // 13194.3f;  //D11
            double riskFree = (obj.RiskFreeRate);//1.8f;   //%  D14
            double historicalMarketReturn = (obj.HistoricMarket);///10.58f;  //% D15 
            double historicalRiskFree = (obj.HistoricRiskReturn);   ///5.01f;  //% D16
            double smallStock = (obj.SmallStock);  //0.001f; //D17
            double rawBeta = obj.RawBeta; ///1.65f;  //D18

            double Dp = obj.PreferredDividend; //0.01f;  //D24
            double Pp = obj.PreferredShare;//1.0f;  //D25
            double marginalTax = obj.TaxRate; //32.00f; //D40
            double projectRiskAdjustment = obj.ProjectRisk;//0.00f; // D41
            int? methodType = obj.MethodType;
            double costOfDebtMethod = 0;
            ////Based On method one or two 
            if (methodType == 1)
            {
                Method1 results = JsonConvert.DeserializeObject<Method1>(obj.Method);
                costOfDebtMethod = (results.RiskFreeRate + results.DefaultSpread);  //D31
            }
            else if (methodType == 2)
            {
                Method2 results = JsonConvert.DeserializeObject<Method2>(obj.Method);
                costOfDebtMethod = (((results.YeildToMaturity / 100) - ((results.ProbabilityOfDefault / 100) * (results.ExpectedLossRate / 100))) * 100); // D37
            }

            double sumOfInputs = marketValueCommon + totalValuePS + marketValueOfND;
            double AdjustedeBeta = ((0.3333) + (0.6667 * (rawBeta)));  //D19
            double marketRiskPremium = ((historicalMarketReturn - historicalRiskFree) + smallStock); //D20
            double costofequity = ((riskFree) + (AdjustedeBeta * marketRiskPremium)); // D21
            double costofpreferredEquity = (Dp / Pp); //D26

            double unleveredcostofcaptial = ((costofequity * (marketValueCommon / (sumOfInputs))) + ((costofpreferredEquity * 100) * (totalValuePS / (sumOfInputs))) + (costOfDebtMethod * (marketValueOfND / (sumOfInputs)))); //D52
            double weightedAverage = (costofequity * (marketValueCommon / (sumOfInputs))) + ((costofpreferredEquity * 100) * (totalValuePS / (sumOfInputs))) + (costOfDebtMethod * (marketValueOfND / (sumOfInputs)) * (1 - (marginalTax / 100))); //D53

            double AdjustedWACC = projectRiskAdjustment + weightedAverage; //D54

            Dictionary<string, double> result = new Dictionary<string, double>
            {
                { "AdjustedeBeta", AdjustedeBeta },
                { "MarketRiskPremium", marketRiskPremium },
                { "Costofequity", costofequity },
                { "CostofpreferredEquity", (costofpreferredEquity*100) },
                { "CostOfDebtMethod", costOfDebtMethod },
                { "Unleveredcostofcaptial", unleveredcostofcaptial },
                { "WeightedAverage", weightedAverage },
                { "AdjustedWACC", AdjustedWACC }
            };

            return result;
        }

        public class Method1
        {
            public double RiskFreeRate;
            public double DefaultSpread;
        }

        public class Method2
        {
            public double YeildToMaturity;
            public double ProbabilityOfDefault;
            public double ExpectedLossRate;
        }
    }
}
