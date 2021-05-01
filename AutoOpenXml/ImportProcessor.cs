using System;
using System.Collections.Generic;
using System.Reflection;
using AutoOpenXml.Exceptions;
using AutoOpenXml.Extensions;
using AutoOpenXml.Models;
using ClosedXML.Excel;

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
            while (NextRow())
            {
                var newRow = new T();
                foreach (var column in Columns)
                {
                    var value = TryReadColumnInfo(column.Index);
                    SetColumnValue(newRow, column, value);
                }

                ImportedData.Add(newRow);
            }
        }

        private void SetColumnValue(T rowData, ColumnInfo prop, CellRead value)
        {
            if (prop.Type == typeof(string))
                ParseToStringValue(rowData, prop, value);
            else if (prop.Type == typeof(bool) || prop.Type == typeof(bool?))
                ParseToBoolValue(rowData, prop, value);
            else if (prop.Type == typeof(int) || prop.Type == typeof(int?))
                ParseToIntValue(rowData, prop, value);
            else if (prop.Type == typeof(decimal) || prop.Type == typeof(decimal?))
                ParseToDecimalValue(rowData, prop, value);
            else if (prop.Type == typeof(DateTime) || prop.Type == typeof(DateTime?))
                ParseToDateTimeValue(rowData, prop, value);
        }

        private void ParseToDateTimeValue(T rowData, ColumnInfo prop, CellRead value)
        {
            if (value.Value == null)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, null);
            else if (value.Type == XLDataType.DateTime)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, (DateTime) value.Value);
            else if (value.Type == XLDataType.Text && !string.IsNullOrEmpty(StringDateFormat))
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, TryParseDateTimeFromStringValue(value.Value.ToString()));
            else
                ThrowImportColumnParseException(prop, value);
        }

        private DateTime TryParseDateTimeFromStringValue(string stringValue)
        {
            try
            {
                var dateValue = DateTime.ParseExact(stringValue,
                    StringDateFormat, System.Globalization.CultureInfo.InvariantCulture);

                return dateValue;
            } catch (FormatException ex)
            {
                throw new ImportColumnParseException($@"Failure on parse '{stringValue}' to type 'DateTime' 
                    With stringDateFormat '{StringDateFormat}'", ex);
            }
            
        }

        private void ParseToStringValue(T rowData, ColumnInfo prop, CellRead value)
        {
            string valueToSet = value.Value != null ? value.Value.ToString() : "";
            
            rowData.GetType().GetProperty(prop.Name).SetValue(rowData, valueToSet);
        }

        private void ParseToDecimalValue(T rowData, ColumnInfo prop, CellRead value)
        {
            if (value.Value == null)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, null);
            else if (value.Type == XLDataType.Text)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, decimal.Parse((string)value.Value));
            else if (value.Type == XLDataType.Number)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, Convert.ToDecimal(value.Value));
            else
                ThrowImportColumnParseException(prop, value);
        }
        
        private void ParseToIntValue(T rowData, ColumnInfo prop, CellRead value)
        {
            if(value.Value == null)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, null);
            else if (value.Type == XLDataType.Text)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, int.Parse((string)value.Value));
            else if (value.Type == XLDataType.Number)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, Convert.ToInt32(value.Value));
            else
                ThrowImportColumnParseException(prop, value);
        }

        private void ParseToBoolValue(T rowData, ColumnInfo prop, CellRead value)
        {
            if (value.Value == null)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, null);
            else if (value.Type == XLDataType.Text)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, bool.Parse((string) value.Value));
            else if (value.Type == XLDataType.Number)
                rowData.GetType().GetProperty(prop.Name).SetValue(rowData, Convert.ToBoolean(value.Value));
            else
                ThrowImportColumnParseException(prop, value);
        }

        private void ThrowImportColumnParseException(ColumnInfo prop, CellRead value)
        {
            throw new ImportColumnParseException($@"Failure on parse '{value.Value}' to type '{prop.Type}' 
                    from type {value.Type.GetName()}");
        }
    }
}