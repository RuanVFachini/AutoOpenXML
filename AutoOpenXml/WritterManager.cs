using ClosedXML.Excel;
using System;
using System.Drawing;

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
                SetCellValue(column.Index, column.Label, TypesEnum.String);

                SetCellBackgroundColor(column.Index, column.HeaderBackgroundColor);
            }

            CurrentRowIndex++;
        }

        private void WriteContentLines()
        {
            foreach (var rowData in DataToExport)
            {
                foreach (var column in Columns)
                {
                    rowData.TryGetValue(column.Label, out var value);
                    SetCellValue(column.Index, value.Value, value.Type, column.Mask);
                }
                CurrentRowIndex++;
            }
        }
        private void SetCellValue(int columnIndex, Object value, TypesEnum type, string mask = null)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, columnIndex);

            if (type == TypesEnum.String)
                cell.DataType = XLDataType.Text;

            if (type == TypesEnum.Bool || type == TypesEnum.NullableBool)
                cell.DataType = XLDataType.Boolean;
                
            if (type == TypesEnum.Int || type == TypesEnum.NullableInt)
                cell.DataType = XLDataType.Number;
            
            if (type == TypesEnum.Long || type == TypesEnum.NullableLong)
                cell.DataType = XLDataType.Number;

            if (type == TypesEnum.Decimal || type == TypesEnum.NullableDecimal)
                cell.DataType = XLDataType.Number;

            if (type == TypesEnum.DateTime || type == TypesEnum.NullableDateTime)
                cell.DataType = XLDataType.DateTime;

            if (!string.IsNullOrEmpty(mask))
                cell.Style.NumberFormat.Format = mask;

            cell.Value = value;
        }

        private void SetCellBackgroundColor(int columnIndex, Color color)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, columnIndex);

            cell.Style.Fill.BackgroundColor = XLColor.FromColor(color);
        }
    }
}