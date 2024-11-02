using ServiceStack;
using Sowfin.Data.Common.Enum;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Sowfin.API.Lib
{
    public class UnitConversion
    {
        public static double[] ConvertUnits(double[] values, string[] units, int flag)
        {
            double[] convertedArray = new double[values.Length];

            for (int i = 0; i < values.Length; i++)
            {
                if(units[i]!=null)
                if (flag == 0)
                {
                    string lowerStr = units[i].ToLower();
                    convertedArray[i] = ReturnBaseValue(lowerStr, values[i]);
                }
                else if (flag == 1)
                {
                    string lowerStr = units[i].ToLower();
                    convertedArray[i] = ReturnUnitValue(lowerStr, values[i]);
                }

            }
            return convertedArray;
        }

        static double ReturnBaseValue(string lowerStr, object valToConvert)
        {
            double convertedVal = 0;
            switch (lowerStr)
            {
                case string str when (lowerStr == "" || lowerStr == ("$")):
                    convertedVal = ParseDouble(valToConvert);
                    break;
                case string str when (lowerStr == "$k" || lowerStr == ("k")):
                    convertedVal = ParseDouble(valToConvert) * 1000;
                    break;
                case string str when (lowerStr == "$m" || lowerStr == ("m")):
                    convertedVal = ParseDouble(valToConvert) * 1000000; // 10^6
                    break;
                case string str when (lowerStr == "$b" || lowerStr == ("b")):
                    convertedVal = ParseDouble(valToConvert) * 1000000000; // 10^9
                    break;
                case string str when (lowerStr == "$t" || lowerStr == ("t")):
                    convertedVal = ParseDouble(valToConvert) * 1000000000000; // 10^12
                    break;
                case string str when (lowerStr == "$p" || lowerStr == ("p")):
                    convertedVal = ParseDouble(valToConvert) * 1000000000000000; // 10^15
                    break;
                default:
                    Console.WriteLine("Default case");
                    break;
            }
            return convertedVal;

        }
        static double ReturnUnitValue(string lowerStr, object valToConvert)
        {
            double convertedVal = 0;
            switch (lowerStr)
            {
                case string str when (lowerStr == "" || lowerStr == ("$")):
                    convertedVal = ParseDouble(valToConvert);
                    break;
                case string str when (lowerStr == "$k" || lowerStr == ("k")):
                    convertedVal = ParseDouble(valToConvert) / 1000;
                    break;
                case string str when (lowerStr == "$m" || lowerStr == ("m")):
                    convertedVal = ParseDouble(valToConvert) / 1000000; // 10^6
                    break;
                case string str when (lowerStr == "$b" || lowerStr == ("b")):
                    convertedVal = ParseDouble(valToConvert) / 1000000000; // 10^9
                    break;
                case string str when (lowerStr == "$t" || lowerStr == ("t")):
                    convertedVal = ParseDouble(valToConvert) / 1000000000000; // 10^12
                    break;
                case string str when (lowerStr == "$p" || lowerStr == ("p")):
                    convertedVal = ParseDouble(valToConvert) / 1000000000000000; // 10^15
                    break;
                default:
                    Console.WriteLine("Default case");
                    break;
            }
            return convertedVal;

        }

        public static double[] ConvertOutputUnits(out string[] outputUnit, params string[] values)
        {
            double[] array = new double[values.Length];
            string[] unitArray = new string[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                double convertedVal = 0;
                string outStr = "";
                if (values[i][0] == '-')
                {
                    convertedVal = ConvertSingleOutput(values[i].Skip(1).Join(""), out outStr) * (-1);
                }
                else
                {
                    convertedVal = ConvertSingleOutput(values[i], out outStr);
                }
                array[i] = convertedVal;
                unitArray[i] = outStr;
            }
            outputUnit = unitArray;
            return array;
        }

        static double ConvertSingleOutput(string value, out string outputUnit)
        {
            double convertValue = 0;
            string unit = null;
            if (value[0] == '$' && Char.IsLetter(value[value.Length - 1]))
            {
                convertValue = ReturnBaseValue(value.Last().ToString().ToLower(), value.Skip(1).Take(value.Length - 2).Join(""));
                unit = value.Last().ToString().ToLower();
            }
            else if (value[0] == '$')
            {
                convertValue = ReturnBaseValue("", value.Skip(1).Join(""));
                unit = "";
            }
            else if (Char.IsLetter(value[value.Length - 1]))
            {
                convertValue = ReturnBaseValue(value.Last().ToString().ToLower(), value.Take(value.Length - 1).Join(""));
                unit = value.Last().ToString().ToLower();
            }
            else
            {
                convertValue = Convert.ToDouble(value);
                unit = "";
            }
            outputUnit = unit;
            return convertValue;
        }

        public static string FindFomartLetter(double value)
        {
            string assignedUnit = null;
            string str = Convert.ToString(value);
            if (str.Contains("."))
            {
                str = str.Split(".")[0];
            }
            int count = str.Length - 1;
            switch (count)
            {
                case int n when count >= 15:
                    assignedUnit = "T";
                    break;

                case int n when count >= 12:
                    assignedUnit = "B";
                    break;

                case int n when (count >= 9 || count >= 7 || count >= 8):
                    assignedUnit = "M";
                    break;

                case int n when count >= 6:
                    assignedUnit = "K";
                    break;

                case int n when count <= 6:
                    assignedUnit = "";
                    break;

            }
            return assignedUnit;

        }

        public static string AssignUnitCurrency(double value)
        {
            string assignedUnit = null;
            string str = Convert.ToString(value);
            if (str.Contains("."))
            {
                str = str.Split(".")[0];
            }
            int count = str.Length - 1;
            switch (count)
            {
                case int n when count >= 15:
                    assignedUnit = AssignCurrency(value / 1000000000000) + "T";
                    break;

                case int n when count >= 12:
                    assignedUnit = AssignCurrency(value / 1000000000) + "B";
                    break;

                case int n when count >= 9 || count >= 7 || count >= 8:
                    assignedUnit = AssignCurrency(value / 1000000) + "M";
                    break;

                case int n when count >= 6:
                    assignedUnit = AssignCurrency(value / 1000) + "K";
                    break;

                case int n when count <= 6:
                    assignedUnit = AssignCurrency(value);
                    break;

            }
            return assignedUnit;
        }

        public static string AssignUnit(double value)
        {
            string assignedUnit = null;
            string str = Convert.ToString(value);
            if (str.Contains("."))
            {
                str = str.Split(".")[0];
            }
            int count = str.Length - 1;
            switch (count)
            {
                case int n when count >= 15:
                    assignedUnit = Convert.ToString(value / 1000000000000) + "T";
                    break;

                case int n when count >= 12:
                    assignedUnit = Convert.ToString(value / 1000000000) + "B";
                    break;

                case int n when count >= 9 || count >= 7 || count >= 8:
                    assignedUnit = Convert.ToString(value / 1000000) + "M";
                    break;

                case int n when count >= 6:
                    assignedUnit = Convert.ToString(value / 1000) + "K";
                    break;

                case int n when count <= 6:
                    assignedUnit = Convert.ToString(value);
                    break;

            }
            return assignedUnit;
        }

        public static double ReturnDividend(string output)
        {
            double dividend = 1;
            string lowerStr = output.ToLower();

            switch (lowerStr)
            {
                case string str when (lowerStr == "" || lowerStr == ("$")):
                    dividend = 1;
                    break;
                case string str when (lowerStr == "$k" || lowerStr == ("k")):
                    dividend = 1000;
                    break;
                case string str when (lowerStr == "$m" || lowerStr == ("m")):
                    dividend = 1000000; // 10^6
                    break;
                case string str when (lowerStr == "$b" || lowerStr == ("b")):
                    dividend = 1000000000; // 10^9
                    break;
                case string str when (lowerStr == "$t" || lowerStr == ("t")):
                    dividend = 1000000000000; // 10^12
                    break;
                case string str when (lowerStr == "$p" || lowerStr == ("p")):
                    dividend = 1000000000000000; // 10^15
                    break;
                default:
                    dividend = 1;
                    break;
            }
            return dividend;

        }

        public static string ReturnCellFormat(string output)
        {
            string str;
            string lowerStr = output.ToLower();
            switch (lowerStr)
            {
                case "$k":
                    str = "$#.00" + "\" K\"";
                    break;
                case "k":
                    str = "#.00" + "\" K\"";
                    break;
                case "$m":
                    str = "$#,,.00" + "\" M\"";
                    break;
                case "m":
                    str = "#,,.00" + "\" M\"";
                    break;
                case "b":
                    str = "#,,,.00" + "\" B\"";
                    break;
                case "$b":
                    str = "$#,,,.00" + "\" B\"";
                    break;
                case "$t":
                    str = "$#,,,,.00" + "\" T\"";
                    break;
                case "t":
                    str = "#,,,,.00" + "\" T\"";
                    break;
                default:
                    str = "#0.00";
                    break;
            }
            return str;
        }



        static string AssignCurrency(double value)
        {
            string str = null;
            if (value < 0)
            {
                value = Math.Abs(value);
                str = "-" + "$" + Convert.ToString(value);

            }
            else
            {
                str = "$" + Convert.ToString(value);
            }

            return str;
        }

        public static Double ParseDouble(Object obj)
        {
            if ((obj == null) || (obj.ToString() == ""))
            {
                return 0;
            }


            return Convert.ToDouble(obj);
        }


        public static double getConvertedValueforCurrency(int? unitId, double? Basicvalue)
        {
            double? Value = 0;
            if (unitId != null)
            {
                if (Basicvalue == null || Basicvalue==0)
                    return 0;
                if (unitId == (int)CurrencyValueEnum.Dollar)
                {
                    Value = Basicvalue;
                }
                else if (unitId == (int)CurrencyValueEnum.DollarK)
                {
                    Value = Basicvalue / 1000;//10^3
                }
                else if (unitId == (int)CurrencyValueEnum.DollarM)
                {
                    Value = Basicvalue / 1000000;//10^6
                }
                else if (unitId == (int)CurrencyValueEnum.DollarB)
                {
                    Value = Basicvalue / 1000000000;//10^9
                }
                else if (unitId == (int)CurrencyValueEnum.DollarT)
                {
                    Value = Basicvalue / 1000000000000;//10^12
                }
            }
            else
                Value = Basicvalue;


            return Convert.ToDouble(Value);
        }

        public static double getConvertedValueforNumbers(int? unitId, double? Basicvalue)
        {
            double? Value = 0;
            if (unitId != null)
            {
                if (Basicvalue == null)
                    return 0;
                if (unitId == (int)NumberCountEnum.EachNumber)
                {
                    Value = Basicvalue;
                }
                else if (unitId == (int)NumberCountEnum.EachK)
                {
                    Value = Basicvalue / 1000;//10^3
                }
                else if (unitId == (int)NumberCountEnum.EachM)
                {
                    Value = Basicvalue / 1000000;//10^6
                }
                else if (unitId == (int)NumberCountEnum.EachB)
                {
                    Value = Basicvalue / 1000000000;//10^9
                }
                else if (unitId == (int)NumberCountEnum.EachT)
                {
                    Value = Basicvalue / 1000000000000;//10^12
                }
            }
            else
                Value = Basicvalue;


            return Convert.ToDouble(Value);
        }

        public static double getBasicValueforCurrency(int? unitId, double? value)
        {
            double? BasicValue = 0;
            if (unitId != null)
            {
                if (value == null)
                    return 0;
                if (unitId == (int)CurrencyValueEnum.Dollar)
                {
                    BasicValue = value;
                }
                else if (unitId == (int)CurrencyValueEnum.DollarK)
                {
                    BasicValue = value * 1000;//10^3
                }
                else if (unitId == (int)CurrencyValueEnum.DollarM)
                {
                    BasicValue = value * 1000000;//10^6
                }
                else if (unitId == (int)CurrencyValueEnum.DollarB)
                {
                    BasicValue = value * 1000000000;//10^9
                }
                else if (unitId == (int)CurrencyValueEnum.DollarT)
                {
                    BasicValue = value * 1000000000000;//10^12
                }
            }
            else
                BasicValue = value;
            


            return Convert.ToDouble(BasicValue);
        }

        public static double getBasicValueforNumbers(int? unitId, double? value)
        {
           
            double? BasicValue = 0;
            if (unitId != null)
            {
                if (value == null)
                    return 0;
                if (unitId == (int)NumberCountEnum.EachNumber)
            {
                BasicValue = value;
            }
            else if (unitId == (int)NumberCountEnum.EachK)
            {
                BasicValue = value * 1000;//10^3
            }
            else if (unitId == (int)NumberCountEnum.EachM)
            {
                BasicValue = value * 1000000;//10^6
            }
            else if (unitId == (int)NumberCountEnum.EachB)
            {
                BasicValue = value * 1000000000;//10^9
            }
            else if (unitId == (int)NumberCountEnum.EachT)
            {
                BasicValue = value * 1000000000000;//10^12
            }
            }
            else
            {
                BasicValue = value;
            }
            return Convert.ToDouble(BasicValue);

        }


        //temporary solution
        public static int? getHigherDenominationUnit(int? unit1, int? unit2)
        {
            
            if (unit1 > unit2)
                return unit1;
            else return unit2;

        }


    }
}
