using System.Collections.Generic;
using System.Reflection;
using AutoOpenXml.Models;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class ExportProcessor<T> where T : new()
    {
        internal List<PropertyInfo> Properties { get; set; } = null;
        internal List<T> Data { get; set; } = null;
        internal XLTableTheme TableStyle { get; set; }
        internal List<Dictionary<string, DataCellValue>> DataToExport { get; set; }
        internal IList<ColumnInfo<T>> Columns { get; set; }

        internal void ProcessData()
        {
            DataToExport = new List<Dictionary<string, DataCellValue>>();
            Columns = Properties.GetInfoColumns<T>(OperatinoEnum.Write);

            foreach (var rowData in Data)
            {
                var newLine = new Dictionary<string, DataCellValue>();

                foreach (var prop in Columns)
                {
                    var columnValue = new DataCellValue()
                    {
                        Type = prop.Type,
                        Value = prop.GetValueFunc.Invoke(rowData)
                    };

                    newLine.Add(prop.Label, columnValue);
                }

                DataToExport.Add(newLine);
            }
            
        }
    }
}
