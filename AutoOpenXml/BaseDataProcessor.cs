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
                    var columnValue = new DataCellValue ()
                    {
                        Type = prop.Type,
                        Value = GetColumnValueByType(rowData, prop)
                    };

                    newLine.Add(prop.Label, columnValue);
                }

                DataToExport.Add(newLine);
            }
            
        }

        private Object GetColumnValueByType(T rowData, ColumnInfo prop)
        {
            if (prop.Type == ColumnTypes.Int || prop.Type == ColumnTypes.Decimal)
                return GetColumnValue<Object>(rowData, prop);
                
            
            if (prop.Type == ColumnTypes.DateTime)
                return GetColumnValue<DateTime>(rowData, prop);

            return GetColumnValue<string>(rowData, prop);
        }

        internal R GetColumnValue<R>(
           T rowData,
            ColumnInfo prop)
        {
            return (R) rowData
                .GetType()
                .GetProperty(prop.Name)
                .GetValue(rowData);
        }
    }
}
