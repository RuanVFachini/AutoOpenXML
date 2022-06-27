using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using AutoOpenXml.Exceptions;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    internal static class CommonManager
    {
        internal static XLTableTheme ExtractTableStyle<T>() where T : new()
        {
            var attributes = ExtractClassAttributes<T>();
            if (attributes.ConstructorArguments.Count > 1)
            {
                var name = (string) attributes.ConstructorArguments[1].Value;
                return XLTableTheme.FromName(name);
            }
                
            return null;
        }

        internal static string ExtractWorksheetName<T>() where T : new()
        {
            var attributes = ExtractClassAttributes<T>();
            return (string) attributes.ConstructorArguments[0].Value;
        }

        private static CustomAttributeData ExtractClassAttributes<T>() where T : new()
        {
                var validClass = (new T())
                    .GetType()
                .CustomAttributes
                .FirstOrDefault(x => x.AttributeType == typeof(ExportWorkSheetAttribute));

            if (validClass != null) return validClass;

            throw new MissingExportWorkSheetAttributeException();
        }

        internal static List<PropertyInfo> ExtractReferenceMapedProperties<T>() where T : new()
        {
            var properties = (new T())
                .GetType()
                .GetProperties()
                .Where(X => X.HasValidAttributeField())
                .ToList();

            if (properties != null) return properties;

            throw new MissingExportColumnAttributeException();
        }
    }
}
