using System;
using System.ComponentModel;

namespace ReportRefresher.Enums
{
    public static class ExtensionMethods
    {
        public static string GetEnumDescription(this Enum enumValue)
        {
            var field = enumValue.GetType().GetField(enumValue.ToString());
            if (Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) is DescriptionAttribute attribute)
            {
                return attribute.Description;
            }
            else
            {
                return enumValue.ToString();
            }
        }
    }
}