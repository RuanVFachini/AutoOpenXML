using System;
using AutoOpenXml.Models;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Fasterflect;

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

        internal static IList<ColumnInfo<T>> GetInfoColumns<T>(this IList<PropertyInfo> properties, OperatinoEnum operatino)
        {
            var result = new List<ColumnInfo<T>>();
            int columnIndex = 1;
            foreach (var prop in properties)
            {
                columnIndex = GetPropIndex(operatino, prop, columnIndex);

                var instance = Expression.Parameter(prop.DeclaringType, "i");
                var property = Expression.Property(instance, prop);
                var convert = Expression.TypeAs(property, typeof(object));
                var propertyFunc = (Func<T, object>)Expression.Lambda(convert, instance).Compile();
                
                result.Add(new ColumnInfo<T>() 
                {
                    Label = Generics.GetColumnCustomProperty<string>(prop, 0),
                    Index = columnIndex,
                    Type = GetTypeEnum(prop.PropertyType),
                    Name = prop.Name,
                    GetValueFunc = propertyFunc,
                    SetValueFunc = Reflect.Setter(typeof(T), prop.Name),
                    Mask = Generics.GetColumnCustomProperty<string>(prop, 2),
                    HeaderBackgroundColor = GetPropHeaderBackgroundColor(prop)
                });;

                columnIndex++;
            }
            return result;
        }

        private static TypesEnum GetTypeEnum(Type propPropertyType)
        {
            if (typeof(string) == propPropertyType)
                return TypesEnum.String;
            if (typeof(bool) == propPropertyType)
                return TypesEnum.Bool;
            if (typeof(bool?) == propPropertyType)
                return TypesEnum.NullableBool;
            if (typeof(int) == propPropertyType)
                return TypesEnum.Int;
            if (typeof(int?) == propPropertyType)
                return TypesEnum.NullableInt;
            if (typeof(long) == propPropertyType)
                return TypesEnum.Long;
            if (typeof(long?) == propPropertyType)
                return TypesEnum.NullableLong;
            if (typeof(Decimal) == propPropertyType)
                return TypesEnum.Decimal;
            if (typeof(Decimal?) == propPropertyType)
                return TypesEnum.NullableDecimal;
            if (typeof(DateTime) == propPropertyType)
                return TypesEnum.DateTime;
            if (typeof(DateTime?) == propPropertyType)
                return TypesEnum.NullableDateTime;

            return TypesEnum.None;
        }

        private static int GetPropIndex(OperatinoEnum operatino, PropertyInfo prop, int columnIndex)
        {
            var proColumnIndex = Generics.GetColumnCustomProperty<int>(prop, 1);
            return operatino == OperatinoEnum.Write ? 
                (proColumnIndex > -1 ? proColumnIndex : columnIndex) : 
                Generics.GetColumnCustomProperty<int>(prop, 1);
        }

        private static Color GetPropHeaderBackgroundColor(PropertyInfo prop)
        {
            var attribute = Generics.GetExportColumnHeaderBackgoundColorAttribute(prop);

            if (attribute == null)
                return Color.Transparent;

            return attribute.HeaderBackgroundColor;
        }
    }
}
