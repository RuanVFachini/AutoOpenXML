using System;
using System.Collections.Generic;
using System.Reflection;
using AutoOpenXml.Exceptions;
using AutoOpenXml.Models;

namespace AutoOpenXml
{
    public class ImportProcessor<T> : ReaderManager<T> where T : new()
    {
        internal List<PropertyInfo> Properties { get; set; } = null;
        internal List<T> ImportedData { get; set; }
        internal IList<ColumnInfo> Columns { get; set; }
        internal bool IsValidDataCell { get; set; }
        internal string StringDateFormat { get; set; }

        internal void ProcessColumns()
        {
            Columns = Properties.GetInfoColumns(OperatinoEnum.Read);
        }

        internal void InitReadData()
        {
            while (IsValidDataCell)
            {
                NextRow();
                var newRow = new T();
                bool someValue = false;
                
                foreach (var column in Columns)
                {
                    var stringColumnValue = TryReadColumnInfo(column.Index);
                    TrySetColumnValue(newRow, column, stringColumnValue);

                    if (!someValue) someValue = stringColumnValue.Length > 0;
                }

                ImportedData.Add(newRow);
                IsValidDataCell = someValue;
            }
        }

        private void TrySetColumnValue(T newRow, ColumnInfo column, string stringColumnValue)
        {
            try
            {
                SetColumnValue(newRow, column, stringColumnValue);
            }
            catch(Exception ex)
            {
                throw new ImportColumnParseException($"Failure on parse '{stringColumnValue}' to type '{column.Type}'", ex);
            }
        }

        private void SetColumnValue(T rowData, ColumnInfo prop, string value)
        {
            if (prop.Type == typeof(int))
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, int.Parse(value));
            else if (prop.Type == typeof(decimal))
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, decimal.Parse(value));
            else if (prop.Type == typeof(string))
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, value);
            else if (prop.Type == typeof(DateTime))
            {
                rowData.GetType().GetProperty(prop.Name)
                    .SetValue(rowData, DateTime.ParseExact(value, StringDateFormat, System.Globalization.CultureInfo.InvariantCulture));
            }
                
        }
    }
}