using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using AutoOpenXml.Models;
using AutoOpenXml.Utils;

namespace AutoOpenXml
{
    public class BaseDataProcessor<T> : BaseManager<T> where T : new()
    {
        internal void SortFields()
        {
            Properties.Sort((a,b) => AttributeUtils.ComparePropertyInfo(a,b));
        }

        internal void ProcessData()
        {
            DataToExport = new List<Dictionary<string, DataCellValue>>();
            Columns = Properties.GetInfoColumns();

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

        internal string GetColumnValue(
           T rowData,
            ColumnInfo prop)
        {
            return rowData
                .GetType()
                .GetProperty(prop.Name)
                .GetValue(rowData)
                .ToString();
        }
    }
}
