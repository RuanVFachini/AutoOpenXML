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

        internal static IList<ColumnInfo> GetInfoColumns(this IList<PropertyInfo> properties, OperatinoEnum operatino)
        {
            var result = new List<ColumnInfo>();
            int columnIndex = 1;
            foreach (var prop in properties)
            {
                columnIndex = GetPropIndex(operatino, prop, columnIndex);

                result.Add(new ColumnInfo() 
                {
                    Label = Generics.GetColumnCustomProperty<string>(prop, 0),
                    Index = columnIndex,
                    Type = prop.PropertyType,
                    Name = prop.Name
                });

                columnIndex++;
            }
            return result;
        }

        private static int GetPropIndex(OperatinoEnum operatino, PropertyInfo prop, int columnIndex)
        {
            var proColumnIndex = Generics.GetColumnCustomProperty<int>(prop, 1);
            return operatino == OperatinoEnum.Write ? 
                (proColumnIndex > -1 ? proColumnIndex : columnIndex) : 
                Generics.GetColumnCustomProperty<int>(prop, 1);
        }
    }
}
