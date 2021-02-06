using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class BaseManager<T>
    {
        internal IXLWorkbook Workbook { get; set; }
        internal IXLWorksheet ActiveWorksheet { get; set; }
        internal List<PropertyInfo> Properties { get; set; } = null;
        internal List<T> Data { get; set; } = null;
        internal List<Dictionary<string, string>> DataToExport { get; set; }
    }
}
