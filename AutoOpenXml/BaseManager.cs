using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using AutoOpenXml.Models;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class BaseManager<T>
    {
        internal IXLWorkbook Workbook { get; set; }
        internal IXLWorksheet ActiveWorksheet { get; set; }
        internal List<PropertyInfo> Properties { get; set; } = null;
        internal List<T> Data { get; set; } = null;
        internal List<Dictionary<string, DataCellValue>> DataToExport { get; set; }
        internal IList<ColumnInfo> Columns { get; set; }
    }
}
