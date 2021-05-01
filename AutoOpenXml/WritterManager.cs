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
            WriteHeaders();
            WriteContentLines();
        }

        internal void WriteHeaders()
        {
            foreach (var column in Columns)
            {
                SetCellValue(column.Index, column.Label, typeof(string));

                SetCellBackgroundColor(column.Index, column.HeaderBackgroundColor);
            }

            CurrentRowIndex++;
        }

        internal void WriteContentLines()
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
        private void SetCellValue(int columnIndex, Object value, Type type, string mask = null)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, columnIndex);

            if (type == typeof(string))
                cell.DataType = XLDataType.Text;

            if (type == typeof(bool) || type == typeof(bool?))
                cell.DataType = XLDataType.Boolean;

            if (type == typeof(int) || type == typeof(int?))
                cell.DataType = XLDataType.Number;

            if (type == typeof(decimal) || type == typeof(decimal?))
                cell.DataType = XLDataType.Number;

            if (type == typeof(DateTime) || type == typeof(DateTime?))
                cell.DataType = XLDataType.DateTime;

            if (mask != null && mask != "")
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