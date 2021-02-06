using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AutoOpenXml
{
    internal static class AttributeExtensions
    {
        internal static bool HasValidAttributeField(this PropertyInfo propertyInfo)
        {
            return propertyInfo.CustomAttributes.Any(x => x.AttributeType == typeof(ExportColumnAttribute));
        }

        internal static int GetFistColumnIndex(this PropertyInfo propertyInfo)
        {
            var attribute = propertyInfo.CustomAttributes.First(x => x.AttributeType == typeof(ExportColumnAttribute));
            return (int) attribute.ConstructorArguments[1].Value;
        }
    }
}
