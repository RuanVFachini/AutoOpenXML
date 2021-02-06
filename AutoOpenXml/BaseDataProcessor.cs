using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
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
            DataToExport = new List<Dictionary<string, string>>();

            foreach (var rowData in Data)
            {
                var newLine = new Dictionary<string, string>();

                foreach (var prop in Properties)
                {
                    var columnName = GetColumnName(prop);
                    var columnValue = GetColumnValue(rowData, prop);

                    newLine.Add(columnName, columnValue);
                }

                DataToExport.Add(newLine);
            }
            
        }

        internal string GetColumnValue(
           T rowData,
            PropertyInfo prop)
        {
            return (string) rowData
                .GetType()
                .GetProperty(prop.Name)
                .GetValue(rowData);
        }

        internal string GetColumnName(PropertyInfo prop)
        {
            return (string) prop.CustomAttributes
                .First(x => x.AttributeType == typeof(ExportColumnAttribute))
                .ConstructorArguments[0]
                .Value;
        }
    }
}
