using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AutoOpenXml
{
    internal static class Generics
    {
        internal static T GetColumnCustomProperty<T>(PropertyInfo prop, int index)
        {
            return (T) prop.CustomAttributes
                .First(x => x.AttributeType == typeof(ExportColumnAttribute))
                .ConstructorArguments[index]
                .Value;
        }
    }
}
