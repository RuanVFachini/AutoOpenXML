using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using AutoOpenXml.Models;

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

        internal static IList<ColumnInfo> GetInfoColumns(this IList<PropertyInfo> properties)
        {
            var result = new List<ColumnInfo>();
            int columnIndex = 1;
            foreach (var prop in properties)
            {
                result.Add(new ColumnInfo() 
                {
                    Label = Generics.GetColumnCustomProperty<string>(prop, 0),
                    Index = columnIndex++,
                    Type = Generics.GetColumnCustomProperty<ColumnTypes>(prop, 2),
                    Name = prop.Name,
                    Func = Generics.GetColumnCustomProperty<Func<Object, Object>>(prop, 3),
                });
            }
            return result;
        }

        
    }
}
