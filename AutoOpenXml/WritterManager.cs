using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

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
                    SetCellValue(column.Index, value.Value, value.Type);
                }
                CurrentRowIndex++;
            }
        }
        private void SetCellValue(int columnIndex, Object value, Type type)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, columnIndex);

            if (type == typeof(string))
            {
                cell.DataType = XLDataType.Text;
                cell.Value = (string) value;
            }
                

            if (type == typeof(int))
            {
                cell.Value = (int) value;
                cell.DataType = XLDataType.Number;
            }

            if (type == typeof(decimal))
            {
                cell.Value = (decimal) value;
                cell.DataType = XLDataType.Number;
            }


            if (type == typeof(DateTime))
            {
                cell.Value = ((DateTime) value);
                cell.DataType = XLDataType.DateTime;
            }

        }
    }
}
