using System.Collections.Generic;
using System.Reflection;
using AutoOpenXml.Utils;

namespace AutoOpenXml.Extensions
{
    internal static class PropertiesExtensions
    {

        internal static List<PropertyInfo> SortFields(this List<PropertyInfo> properties)
        {
            properties.Sort((a, b) => AttributeUtils.ComparePropertyInfo(a, b));
            return properties;
        }
    }
}
