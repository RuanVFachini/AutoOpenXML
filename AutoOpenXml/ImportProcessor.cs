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
        internal IList<ColumnInfo<T>> Columns { get; set; }
        
        internal bool IsValidDataCell { get; set; }
        internal string StringDateFormat { get; set; }

        internal void ProcessColumns()
        {
            Columns = Properties.GetInfoColumns<T>(OperatinoEnum.Read);
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

        private void SetColumnValue(T rowData, ColumnInfo<T> prop, CellRead value)
        {
            if (prop.Type == TypesEnum.String)
                ParseToStringValue(rowData, prop, value);
            else if (prop.Type == TypesEnum.Bool)
                ParseToBoolValue(rowData, prop, value);
            else if (prop.Type == TypesEnum.NullableBool)
                SetNullableColumnValue(rowData, prop, value, ParseToBoolValue);
            else if (prop.Type == TypesEnum.Int)
                ParseToIntValue(rowData, prop, value);
            else if (prop.Type == TypesEnum.NullableInt)
                SetNullableColumnValue(rowData, prop, value, ParseToIntValue);
            else if (prop.Type == TypesEnum.Long)
                ParseToIntValue(rowData, prop, value);
            else if (prop.Type == TypesEnum.NullableLong)
                SetNullableColumnValue(rowData, prop, value, ParseToLongValue);
            else if (prop.Type == TypesEnum.Decimal)
                ParseToDecimalValue(rowData, prop, value);
            else if (prop.Type == TypesEnum.NullableDecimal)
                SetNullableColumnValue(rowData, prop, value, ParseToDecimalValue);
            else if (prop.Type == TypesEnum.DateTime)
                ParseToDateTimeValue(rowData, prop, value);
            else if (prop.Type == TypesEnum.NullableDateTime)
                SetNullableColumnValue(rowData, prop, value, ParseToDateTimeValue);
        }

        private void SetNullableColumnValue(T rowData, ColumnInfo<T> prop, CellRead value, Action<T,ColumnInfo<T>,CellRead> action)
        {
            if (value.Value == null || (string)value.Value == string.Empty)
                prop.SetValueFunc(rowData, null);
            else
                action.Invoke(rowData, prop, value);
        }

        private void ParseToDateTimeValue(T rowData, ColumnInfo<T> prop, CellRead value)
        {
            if (value.Value == null)
                return;
            
            if (prop.Type == TypesEnum.String && ((string) value.Value).Trim() == "")
                prop.SetValueFunc(rowData, null);
            else if (value.Type == XLDataType.DateTime)
                prop.SetValueFunc(rowData, (DateTime) value.Value);
            else if (value.Type == XLDataType.Text && !string.IsNullOrEmpty(StringDateFormat))
                prop.SetValueFunc(rowData, TryParseDateTimeFromStringValue(value.Value.ToString()));
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

        private void ParseToStringValue(T rowData, ColumnInfo<T> prop, CellRead value)
        {
            string valueToSet = value.Value != null ? value.Value.ToString() : "";
            
            prop.SetValueFunc(rowData, valueToSet);
        }

        private void ParseToDecimalValue(T rowData, ColumnInfo<T> prop, CellRead value)
        {
            if (IsEmptyCell(value, prop))
                prop.SetValueFunc(rowData, null);
            else if (value.Type == XLDataType.Text)
                prop.SetValueFunc(rowData, decimal.Parse((string)value.Value));
            else if (value.Type == XLDataType.Number)
                prop.SetValueFunc(rowData, Convert.ToDecimal(value.Value));
            else
                ThrowImportColumnParseException(prop, value);
        }
        
        private void ParseToIntValue(T rowData, ColumnInfo<T> prop, CellRead value)
        {
            if(IsEmptyCell(value, prop))
                prop.SetValueFunc(rowData, null);
            else if (value.Type == XLDataType.Text)
                prop.SetValueFunc(rowData, int.Parse((string)value.Value));
            else if (value.Type == XLDataType.Number)
                prop.SetValueFunc(rowData, Convert.ToInt32(value.Value));
            else
                ThrowImportColumnParseException(prop, value);
        }
        
        private void ParseToLongValue(T rowData, ColumnInfo<T> prop, CellRead value)
        {
            if(IsEmptyCell(value, prop))
                prop.SetValueFunc(rowData, null);
            else if (value.Type == XLDataType.Text)
                prop.SetValueFunc(rowData, long.Parse((string)value.Value));
            else if (value.Type == XLDataType.Number)
                prop.SetValueFunc(rowData, Convert.ToInt64(value.Value));
            else
                ThrowImportColumnParseException(prop, value);
        }

        private void ParseToBoolValue(T rowData, ColumnInfo<T> prop, CellRead value)
        {
            if (IsEmptyCell(value, prop))
                prop.SetValueFunc(rowData, null);
            else if (value.Type == XLDataType.Boolean)
                prop.SetValueFunc(rowData, (bool) value.Value);
            else if (value.Type == XLDataType.Text)
                prop.SetValueFunc(rowData, bool.Parse((string) value.Value));
            else if (value.Type == XLDataType.Number)
                prop.SetValueFunc(rowData, Convert.ToBoolean(value.Value));
            else
                ThrowImportColumnParseException(prop, value);
        }

        private bool IsEmptyCell(CellRead value, ColumnInfo<T> columnInfo)
        {
            return (columnInfo.Type == TypesEnum.String && (string)value.Value == "")
                || value.Value == null;
        }

        private void ThrowImportColumnParseException(ColumnInfo<T> prop, CellRead value)
        {
            throw new ImportColumnParseException($@"Failure on parse '{value.Value}' to type '{prop.Type}' 
                    from type {value.Type.GetName()}");
        }
    }
}