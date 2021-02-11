using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using AutoOpenXml.Exceptions;

namespace AutoOpenXml
{
    internal static class CommonManager
    {
        internal static string ExtractWorksheetName<T>() where T : new()
        {
            var validClass = (new T())
                .GetType()
                .CustomAttributes
                .FirstOrDefault(x => x.AttributeType == typeof(ExportWorkSheetAttribute));

            if (validClass != null) return (string) validClass.ConstructorArguments[0].Value;

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
