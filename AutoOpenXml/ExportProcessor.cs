using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using AutoOpenXml.Models;
using AutoOpenXml.Utils;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;

namespace AutoOpenXml
{
    public class ExportProcessor<T> where T : new()
    {
        internal List<PropertyInfo> Properties { get; set; } = null;
        internal List<T> Data { get; set; } = null;
        internal XLTableTheme TableStyle { get; set; }
        internal List<Dictionary<string, DataCellValue>> DataToExport { get; set; }
        internal IList<ColumnInfo> Columns { get; set; }

        internal void ProcessData()
        {
            DataToExport = new List<Dictionary<string, DataCellValue>>();
            Columns = Properties.GetInfoColumns(OperatinoEnum.Write);

            foreach (var rowData in Data)
            {
                var newLine = new Dictionary<string, DataCellValue>();

                foreach (var prop in Columns)
                {
                    var columnValue = new DataCellValue()
                    {
                        Type = prop.Type,
                        Value = GetColumnValue(rowData, prop)
                    };

                    newLine.Add(prop.Label, columnValue);
                }

                DataToExport.Add(newLine);
            }
            
        }

        private Object GetColumnValue(T rowData, ColumnInfo prop)
        {
            return rowData
                .GetType()
                .GetProperty(prop.Name)
                .GetValue(rowData);
        }
    }
}
