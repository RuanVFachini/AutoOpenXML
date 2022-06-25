using ClosedXML.Excel;

namespace AutoOpenXml.Extensions
{
    internal static class XLDataTypeExtensions
    {
        internal static string GetName(this XLDataType type)
        {
            switch (type)
            {
                case XLDataType.Number: return "Number";
                case XLDataType.Text: return "Text"; 
                case XLDataType.Boolean: return "Boolean"; 
                case XLDataType.DateTime: return "DateTime"; 
                case XLDataType.TimeSpan: return "TimeSpan";
                default: return "Unknown";
            };
        }
    }
}
