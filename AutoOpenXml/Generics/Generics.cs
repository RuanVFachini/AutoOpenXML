using System.Linq;
using System.Reflection;

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

        internal static ExportColumnHeaderBackgoundColorAttribute GetExportColumnHeaderBackgoundColorAttribute(PropertyInfo prop)
        {
            return prop.GetCustomAttributes()
                .FirstOrDefault(x => x is ExportColumnHeaderBackgoundColorAttribute)
                as ExportColumnHeaderBackgoundColorAttribute;
        }
    }
}