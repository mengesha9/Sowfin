using System;
using System.Collections.Generic;

namespace Sowfin.API.Lib
{
    public class MathCapitalBudgeting
    {
        public static Dictionary<string, object> SummaryOutput(Dictionary<string, object> Obj)
        { // For scenrio and sensi check line 442 in calengine
            List<double> volume = (List<double>)Obj["Volume"];
            List<double> unitPrice = (List<double>)Obj["UnitPrice"];
            List<double> unitCost = (List<double>)Obj["UnitCost"];
            //List<double> sellingGeneralAdmin = Object["SGA"];
            //List<double> RD = Object["RA"];  //RD
            List<double> Nwc = (List<double>)Obj["NWC"];
            List<double> Capex = (List<double>)Obj["Capex"];
            List<double> FixedCost = (List<double>)Obj["Fixed"];
            List<double> deprections = (List<double>)Obj["TotalDepreciation"]; ///list directly from output 
            double MarginalTax = Convert.ToDouble(Obj["MarginalTax"]);
            double rWACC = Convert.ToDouble(Obj["DiscountRate"]);

            List<double> sales = new List<double>();
            List<double> cogs = new List<double>();
            List<double> grossMargins = new List<double>();
            List<double> sellingGenerals = new List<double>();
            List<double> rAndDs = new List<double>();
            List<double> operatingIncomes = new List<double>();
            List<double> incomeTaxs = new List<double>();
            List<double> unleveredNetIcomes = new List<double>();
            List<double> nwcs = new List<double>();
            List<double> increaseNwcs = new List<double>();
            List<double> freeCashFlows = new List<double>();
            List<double> leveredValues = new List<double>();
            List<double> discountFactors = new List<double>();
            List<double> discountCashFlow = new List<double>();
            double sale = 0; // (100/100) This is percent 
            double cog = 0;
            double grossMargin = 0;
            double sellingGeneral = 0;
            double rAndD = 0;
            double operatingIncome = 0;
            double incomeTax = 0;
            double unleveredNetIcome = 0;
            double nwc = 0;
            double npv = 0;
            for (int i = 0; i < volume.Count; i++)
            {

                sale = (volume[i] * unitPrice[i]); // (100/100) This is percent 
                cog = (volume[i] * unitCost[i]);
                grossMargin = sale - cog;
                //sellingGeneral = sellingGeneralAdmin[i] * (100 / 100);
                //rAndD = RD[i] * (100 / 100);
                operatingIncome = grossMargin - FixedCost[i] - deprections[i];
                incomeTax = operatingIncome * (MarginalTax / 100);
                unleveredNetIcome = operatingIncome - incomeTax;
                nwc = sale * (Nwc[i] / 100);
                sales.Add(sale);
                cogs.Add(cog);
                grossMargins.Add(grossMargin);
                sellingGenerals.Add(sellingGeneral);
                rAndDs.Add(rAndD);
                operatingIncomes.Add(operatingIncome);
                incomeTaxs.Add(incomeTax);
                unleveredNetIcomes.Add(unleveredNetIcome);
                nwcs.Add(nwc);
            }

            double increaseNwc = 0;
            for (int j = 0; j < nwcs.Count; j++)
            {

                if (j == 0)
                {
                    increaseNwc = 0;
                }
                else
                {
                    increaseNwc = nwcs[j] - nwcs[j - 1];
                }
                double freeCashFlow = unleveredNetIcomes[j] + deprections[j] - Capex[j] - increaseNwc;

                increaseNwcs.Add(increaseNwc);
                freeCashFlows.Add(freeCashFlow);
            }

            int count = freeCashFlows.Count;
            double discountFactor = 0;
            double discountCash = 0;
            leveredValues = Recursive(freeCashFlows, leveredValues, 0, ref count, rWACC);
            leveredValues.Reverse();
            for (int t = 0; t < volume.Count; t++)
            {
                discountFactor = Math.Pow(1 + (rWACC / 100), t);
                discountFactors.Add(discountFactor);
                discountCash = freeCashFlows[t] / discountFactor;
                discountCashFlow.Add(discountCash);
                npv += (discountCash);
            }
            List<double> npvList = new List<double> { };
            npvList.Add(npv);
            Dictionary<string, object> result = new Dictionary<string, object>
            {
                { "sales",sales},
                { "cogs",cogs},
                { "grossMargins",grossMargins},
                { "sellingGenerals",sellingGenerals },
                { "rAndDs",rAndDs},
                { "operatingIncomes",operatingIncomes },
                { "incomeTaxs",incomeTaxs},
                { "unleveredNetIcomes",unleveredNetIcomes},
                { "nwcs",nwcs},
                { "increaseNwcs",increaseNwcs},
                { "freeCashFlows",freeCashFlows},
                { "leveredValues",leveredValues},
                { "discountFactors",discountFactors},
                { "discountCashFlow",discountCashFlow},
                { "npv", npvList }
            };

            return result;
        }

        static List<double> Recursive(List<double> values, List<double> leveredValues, double value, ref int count, double rWACC)
        {
            count--;
            if (count == 0)
            {
                return leveredValues;
            }
            value = ((values[count] + value) / (1 + (rWACC / 100)));
            leveredValues.Add(value);
            return Recursive(values, leveredValues, value, ref count, rWACC);
        }

        // Code Merge

        public static Dictionary<string, object> SummaryOutput_Double(Dictionary<string, List<double>> Obj)
        { // For scenrio and sensi check line 442 in calengine
            List<double> volume = (List<double>)Obj["Volume"];
            List<double> unitPrice = (List<double>)Obj["UnitPrice"];
            List<double> unitCost = (List<double>)Obj["UnitCost"];
            //List<double> sellingGeneralAdmin = Object["SGA"];
            //List<double> RD = Object["RA"];  //RD
            List<double> Nwc = (List<double>)Obj["NWC"];
            List<double> Capex = (List<double>)Obj["Capex"];
            List<double> FixedCost = (List<double>)Obj["Fixed"];
            List<double> deprections = (List<double>)Obj["TotalDepreciation"]; ///list directly from output 
            double MarginalTax = Convert.ToDouble(Obj["MarginalTax"][0]);
            double rWACC = Convert.ToDouble(Obj["DiscountRate"][0]);

            List<double> sales = new List<double>();
            List<double> cogs = new List<double>();
            List<double> grossMargins = new List<double>();
            List<double> sellingGenerals = new List<double>();
            List<double> rAndDs = new List<double>();
            List<double> operatingIncomes = new List<double>();
            List<double> incomeTaxs = new List<double>();
            List<double> unleveredNetIcomes = new List<double>();
            List<double> nwcs = new List<double>();
            List<double> increaseNwcs = new List<double>();
            List<double> freeCashFlows = new List<double>();
            List<double> leveredValues = new List<double>();
            List<double> discountFactors = new List<double>();
            List<double> discountCashFlow = new List<double>();
            double sale = 0; // (100/100) This is percent 
            double cog = 0;
            double grossMargin = 0;
            double sellingGeneral = 0;
            double rAndD = 0;
            double operatingIncome = 0;
            double incomeTax = 0;
            double unleveredNetIcome = 0;
            double nwc = 0;
            double npv = 0;
            for (int i = 0; i < volume.Count; i++)
            {

                sale = (volume[i] * unitPrice[i]); // (100/100) This is percent 
                cog = (volume[i] * unitCost[i]);
                grossMargin = sale - cog;
                //sellingGeneral = sellingGeneralAdmin[i] * (100 / 100);
                //rAndD = RD[i] * (100 / 100);
                operatingIncome = grossMargin - FixedCost[i] - deprections[i];
                incomeTax = operatingIncome * (MarginalTax / 100);
                unleveredNetIcome = operatingIncome - incomeTax;
                nwc = sale * (Nwc[i] / 100);
                sales.Add(sale);
                cogs.Add(cog);
                grossMargins.Add(grossMargin);
                sellingGenerals.Add(sellingGeneral);
                rAndDs.Add(rAndD);
                operatingIncomes.Add(operatingIncome);
                incomeTaxs.Add(incomeTax);
                unleveredNetIcomes.Add(unleveredNetIcome);
                nwcs.Add(nwc);
            }

            double increaseNwc = 0;
            for (int j = 0; j < nwcs.Count; j++)
            {

                if (j == 0)
                {
                    increaseNwc = 0;
                }
                else
                {
                    increaseNwc = nwcs[j] - nwcs[j - 1];
                }
                double freeCashFlow = unleveredNetIcomes[j] + deprections[j] - Capex[j] - increaseNwc;

                increaseNwcs.Add(increaseNwc);
                freeCashFlows.Add(freeCashFlow);
            }

            int count = freeCashFlows.Count;
            double discountFactor = 0;
            double discountCash = 0;
            leveredValues = Recursive(freeCashFlows, leveredValues, 0, ref count, rWACC);
            leveredValues.Reverse();
            for (int t = 0; t < volume.Count; t++)
            {
                discountFactor = Math.Pow(1 + (rWACC / 100), t);
                discountFactors.Add(discountFactor);
                discountCash = freeCashFlows[t] / discountFactor;
                discountCashFlow.Add(discountCash);
                npv += (discountCash);
            }
            List<double> npvList = new List<double> { };
            npvList.Add(npv);
            Dictionary<string, object> result = new Dictionary<string, object>
            {
                { "sales",sales},
                { "cogs",cogs},
                { "grossMargins",grossMargins},
                { "sellingGenerals",sellingGenerals },
                { "rAndDs",rAndDs},
                { "operatingIncomes",operatingIncomes },
                { "incomeTaxs",incomeTaxs},
                { "unleveredNetIcomes",unleveredNetIcomes},
                { "nwcs",nwcs},
                { "increaseNwcs",increaseNwcs},
                { "freeCashFlows",freeCashFlows},
                { "leveredValues",leveredValues},
                { "discountFactors",discountFactors},
                { "discountCashFlow",discountCashFlow},
                { "npv", npvList }
            };

            return result;
        }


    }
}
