using ClosedXML.Excel;
using System;
using System.Drawing;
using AutoOpenXml.Models;

namespace AutoOpenXml
{
    public class WritterManager<T> : ExportProcessor<T> where T : new()
    {
        internal IXLWorksheet ActiveWorksheet { get; set; }
        private int CurrentRowIndex { get; set; } = 1;

        internal void WriteFile()
        {
            CreateTable();
            WriteHeaders();
            FormatColumns();
            WriteContentLines();
        }

        private void CreateTable()
        {
            if (TableStyle == null) return;
            
            var table = ActiveWorksheet.Range(1, 1, Data.Count + 1, Columns.Count).CreateTable();
            table.Theme = TableStyle;

        }

        private void WriteHeaders()
        {
            foreach (var column in Columns)
            {
                var cell = ActiveWorksheet.Cell(CurrentRowIndex, column.Index);
                cell.DataType = XLDataType.Text;
                cell.Value = column.Label;

                SetCellBackgroundColor(column.Index, column.HeaderBackgroundColor);
            }

            CurrentRowIndex++;
        }
        
        internal void FormatColumns()
        {
            foreach (var column in Columns)
            {
                FormatColumn(column);
            }
        }

        private void WriteContentLines()
        {
            foreach (var rowData in DataToExport)
            {
                foreach (var column in Columns)
                {
                    rowData.TryGetValue(column.Label, out var value);
                    var cell = ActiveWorksheet.Cell(CurrentRowIndex, column.Index);
                    cell.Value = value.Value;
                }
                CurrentRowIndex++;
            }
        }
        private void FormatColumn(ColumnInfo<T> column)
        {
            var range = ActiveWorksheet.Range(2, column.Index, Data.Count + 1, column.Index);
            
            if (!string.IsNullOrEmpty(column.Mask))
                range.Style.NumberFormat.SetFormat(column.Mask);
            
            ActiveWorksheet.Column(column.Index).AdjustToContents();
            
            if (column.Type == TypesEnum.String)
                range.DataType = XLDataType.Text;

            if (column.Type == TypesEnum.Bool || column.Type == TypesEnum.NullableBool)
                range.DataType = XLDataType.Boolean;
                
            if (column.Type == TypesEnum.Int || column.Type == TypesEnum.NullableInt)
                range.DataType = XLDataType.Number;
            
            if (column.Type == TypesEnum.Long || column.Type == TypesEnum.NullableLong)
                range.DataType = XLDataType.Number;

            if (column.Type == TypesEnum.Decimal || column.Type == TypesEnum.NullableDecimal)
                range.DataType = XLDataType.Number;

            if (column.Type == TypesEnum.DateTime || column.Type == TypesEnum.NullableDateTime)
                range.DataType = XLDataType.DateTime;
        }

        private void SetCellBackgroundColor(int columnIndex, Color color)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, columnIndex);

            cell.Style.Fill.BackgroundColor = XLColor.FromColor(color);
        }
    }
}