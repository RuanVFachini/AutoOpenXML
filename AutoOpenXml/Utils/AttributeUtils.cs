using System.Reflection;

namespace AutoOpenXml.Utils
{
    internal static class AttributeUtils
    {
        internal static int ComparePropertyInfo(PropertyInfo x, PropertyInfo y)
        {
            var xIndex = x.GetFistColumnIndex();
            var yIndex = y.GetFistColumnIndex();
            return xIndex > yIndex ? 1 : xIndex < yIndex ? -1 : 0;
        }

    }
}
