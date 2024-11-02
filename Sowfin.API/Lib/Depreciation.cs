using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;


namespace Sowfin.API.Lib
{
    public class Depreciation
    {
        public static object[] CalculateDepreciation(double year, int addToYear, object[] capCost, double[] customArray,
            double[][] customPercenList, double[] customAddToyear)
        {
            double years = year;
            int duration = (int)(year);
            List<double> customcapCost = new List<double>();
            List<int> index = new List<int>();
            List<double> capSingle = new List<double>();
            List<List<double>> listOfLists = new List<List<double>>();
            
            if (customArray.Length == 0 && customPercenList.Length == 0)
            {
                double opt = (double)(1.0 / (years));
                for (int i = 0; i < capCost.Length; i++)
                {
                    if ((double)capCost[i] != 0)
                    {
                        double res = ((double)capCost[i]) * (opt);
                        capSingle.Add(res);
                        index.Add(i + addToYear);
                    }
                }
                for (int j = 0; j < index.Count; j++)
                {
                    List<double> resultDep = FindDep(index[j], capSingle[j], capCost.Length, duration, addToYear);
                    listOfLists.Add(resultDep);
                }
            }
            else if (customPercenList.Length == 0)
            {
                List<double> CustomList = new List<double>();
                for (int i = 0; i < capCost.Length; i++)
                {
                    double res = 0;
                    List<double> capSingleCustom = new List<double>();
                    if ((double)capCost[i] != 0)
                    {
                        for (int e = 0; e < customArray.Length; e++)
                        {
                            res = ((double)capCost[i] * customArray[e]) / 100;
                            capSingleCustom.Add(res);
                        }
                        CustomList = FindCustomDep((i + addToYear), capSingleCustom, capCost.Count(), duration, addToYear);
                        listOfLists.Add(CustomList);
                    }
                }
            }
            else
            {
                List<double> CustomList = new List<double>();
                for (int i = 0; i < capCost.Length; i++)
                {
                    if ((double)capCost[i] != 0)
                    {
                        customcapCost.Add((double)capCost[i]);
                        index.Add(i);
                    }
                }
                for (int r = 0; r < customcapCost.Count; r++)
                {
                    if (r < customPercenList.Length && customPercenList[r].Length > 0 && r < customAddToyear.Length)
                    {

                    CustomList = FindFullyDepCost(customcapCost[r], customPercenList[r], (int)(index[r] + customAddToyear[r]), capCost.Length, duration, (int)customAddToyear[r]);
                    listOfLists.Add(CustomList);
                    }
                }
            }

            List<double> resultSum = new List<double>();
            int maxIndex = getMaxValue(listOfLists);
            for (int w = 0; w < maxIndex; w++)
            {
                double colSum = 0;
                for (int e = 0; e < listOfLists.Count; e++)
                {
                    if (w < listOfLists[e].Count) // Check if the index exists in the current list
                        colSum += listOfLists[e][w];
                }
                resultSum.Add(colSum);
            }

            var finalSum = resultSum.ToArray();
            Array.Resize<double>(ref finalSum, capCost.Length);
            object[] object_array = new object[finalSum.Length];
            finalSum.CopyTo(object_array, 0);
            return object_array;
        }

        public static int getMaxValue(List<List<double>> numbers)
        {
            if (numbers == null || numbers.Count == 0 || numbers.Any(n => n == null))
                return 0; // Safeguard against null or empty lists, returning 0 if no valid lists exist

            int maxValue = numbers[0].Count;
            for (int j = 1; j < numbers.Count; j++)
            {
                if (numbers[j].Count > maxValue)
                {
                    maxValue = numbers[j].Count;
                }
            }
            return maxValue;
        }



        // public static object[] CalculateDepreciation(double year, int addToYear, object[] capCost, double[] customArray,
        //      double[][] customPercenList, double[] customAddToyear)
        // {
        //     double years = year;
        //     int duration = (int)(year);
        //     List<double> customcapCost = new List<double>();
        //     List<int> index = new List<int>();
        //     List<double> capSingle = new List<double>();
        //     List<List<double>> listOfLists = new List<List<double>>();
        //     if (customArray.Length == 0 && customPercenList.Length == 0)
        //     {
        //         double opt = (double)(1.0 / (years));
        //         for (int i = 0; i < capCost.Length; i++)
        //         {
        //             if ((double)capCost[i] != 0)
        //             {
        //                 double res = ((double)capCost[i]) * (opt);
        //                 capSingle.Add(res);
        //                 index.Add(i + addToYear);
        //             }
        //         }
        //         for (int j = 0; j < index.Count; j++)
        //         {
        //             List<double> resultDep = FindDep(index[j], capSingle[j], capCost.Length, duration, addToYear);

        //             listOfLists.Add(resultDep);
        //         }
        //     }
        //     else if (customPercenList.Length == 0)
        //     {
        //         List<double> CustomList = new List<double>();
        //         for (int i = 0; i < capCost.Length; i++)
        //         {
        //             double res = 0;
        //             List<double> capSingleCustom = new List<double>();
        //             if ((double)capCost[i] != 0)
        //             {
        //                 for (int e = 0; e < customArray.Length; e++)
        //                 {
        //                     res = ((double)capCost[i] * customArray[e]) / 100;
        //                     capSingleCustom.Add(res);
        //                 }
        //                 CustomList = FindCustomDep((i + addToYear), capSingleCustom, capCost.Count(), duration, addToYear);
        //                 listOfLists.Add(CustomList);
        //             }
        //         }
        //     }
        //     else
        //     {
        //         List<double> CustomList = new List<double>();
        //         for (int i = 0; i < capCost.Length; i++)
        //         {
        //             if ((double)capCost[i] != 0)
        //             {
        //                 customcapCost.Add((double)capCost[i]);
        //                 index.Add(i);
        //             }
        //         }
        //         for (int r = 0; r < customcapCost.Count; r++)
        //         {
        //             CustomList = FindFullyDepCost(customcapCost[r], customPercenList[r], (int)(index[r] + customAddToyear[r]), capCost.Length, duration, (int)customAddToyear[r]);
        //             listOfLists.Add(CustomList);
        //         }


        //     }

        //     List<double> resultSum = new List<double>();
        //     for (int w = 0; w < getMaxValue(listOfLists); w++)
        //     {
        //         double colSum = 0;
        //         for (int e = 0; e < listOfLists.Count; e++)
        //         {
        //             try { colSum += listOfLists[e][w]; }
        //             catch (IndexOutOfRangeException arr)
        //             {
        //                 colSum += 0.0;
        //             }
        //         }
        //         resultSum.Add(colSum);
        //     }

        //     var finalSum = resultSum.ToArray();
        //     Array.Resize<double>(ref finalSum, capCost.Length);
        //     object[] object_array = new object[finalSum.Length];
        //     finalSum.CopyTo(object_array, 0);
        //     return object_array;
        // }
        public static List<double> FindDep(int index, double value, int length, int years, int addToYear)
        {
            int t = 0;
            int u = 0;
            double[] resultDep = new double[length + years + addToYear + 2];
            while (t < resultDep.Length)
            {
                if (t == index)
                {
                    for (int q = 0; q < years; q++)
                    {
                        u = t + q;
                        resultDep[u] = value;
                    }
                    t = u + 1;
                }
                else
                {
                    resultDep[t] = (double)0.0;
                    t++;
                }
            }
            List<double> result = new List<double>(resultDep);
            return result;

        }
        public static List<double> FindCustomDep(int index, List<double> values, int length, int years, int addToYear)
        {
            int t = 0;
            int u = 0;
            double[] resultDep = new double[length + years + addToYear + 2];
            while (t < resultDep.Length)
            {
                if (t == index)
                {
                    for (int q = 0; q < values.Count; q++)
                    {
                        u = t + q;
                        resultDep[u] = values[q];
                    }
                    t = u + 1;
                }
                else
                {
                    resultDep[t] = (double)0.0;
                    t++;
                }
            }
            List<double> result = new List<double>(resultDep);
            return result;
        }

        public static List<double> FindFullyDepCost(double value, double[] percentArray, int index, int length, int years, int addToYear)
        {
            List<double> resultant = new List<double>();
            for (int z = 0; z < percentArray.Length; z++)
            {
                double result = (value * percentArray[z]) / 100;
                resultant.Add(result);
            }
            int t = 0;
            int u = 0;
            double[] resultDep = new double[length + years + addToYear + 2];
            while (t < resultDep.Length)
            {
                if (t == index)
                {
                    for (int q = 0; q < (resultant.Count); q++)
                    {
                        u = t + q;
                        resultDep[u] = resultant[q];
                    }
                    t = u + 1;
                }
                else
                {
                    resultDep[t] = (double)0.0;
                    t++;
                }
            }
            List<double> results = new List<double>(resultDep);
            return results;
        }
        // public static int getMaxValue(List<List<double>> numbers)
        // {
        //      if (numbers == null || numbers.Count == 0){
        //         return 0;
        //      }
    

        //     int maxValue = numbers[0].Count;
        //     for (int j = 0; j < numbers.Count; j++)
        //     {
        //         if (numbers[j].Count > maxValue)
        //         {
        //             maxValue = numbers[j].Count;
        //         }
        //     }
        //     return maxValue;
        // }

    }
}