using Sowfin.Data.Repositories;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Text;

namespace Sowfin.Data.Common.Helper
{
     public static class EnumHelper
    {
        public static string DescriptionAttr<T>(this T source)
        {
            if(source == null)
            {
                return string.Empty;
            }
            // FieldInfo fi = source.GetType().GetField(source.ToString());
            FieldInfo? fi = source?.GetType().GetField(source.ToString() ?? string.Empty);
            if (fi == null)
            {
                return string.Empty; // or return a fallback description
            }
            DescriptionAttribute[] attributes = (DescriptionAttribute[])fi.GetCustomAttributes(
                typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0) return attributes[0].Description;
            else return source.ToString();
        }

        public static List<SelectListItem> GetEnumListbyName<T>()
        {
            Type type = typeof(T);            
            List<SelectListItem>itemsList = new List<SelectListItem>();
            foreach (var e in System.Enum.GetValues(type))
            {
                FieldInfo fi = type.GetField(e.ToString());
                DescriptionAttribute[] attributes = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);
                itemsList.Add(new SelectListItem
                {
                    Text = (attributes!=null && attributes.Length > 0) ? attributes[0].Description : e.ToString(),
                    Value = (int)e

                });
            }
            return itemsList;
        }

        public static List<ValueTextWrapper> GetEnumDescriptions<T>()
        {
            Type enumType = typeof(T);

            if (enumType.BaseType != typeof(System.Enum))
                throw new ArgumentException("T is not System.Enum");

            List<ValueTextWrapper> enumValList = new List<ValueTextWrapper>();
            foreach (var e in System.Enum.GetValues(typeof(T)))
            {
                var fi = e.GetType().GetField(e.ToString());
                var attributes = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

                enumValList.Add(new ValueTextWrapper { value = (int)e, text = (attributes.Length > 0) ? attributes[0].Description : e.ToString() });
            }
            return enumValList;
        }


    }

    public class ValueTextWrapper
    {
        public int value { get; set; }
        public string? text { get; set; }
    }
}
