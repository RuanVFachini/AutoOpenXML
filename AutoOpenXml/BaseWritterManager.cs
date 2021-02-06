using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoOpenXml
{
    public class BaseWritterManager<T> : BaseDataProcessor<T> where T : new()
    {
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
                SetCellValue(column.Index, column.Label);
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
        private void SetCellValue(int columnIndex, string value, ColumnTypes type = ColumnTypes.Text)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, columnIndex);

            cell.Value = value;

            if (type == ColumnTypes.Text)
                cell.DataType = ClosedXML.Excel.XLDataType.Text;

            if (type == ColumnTypes.Number)
                cell.DataType = ClosedXML.Excel.XLDataType.Number;

            if (type == ColumnTypes.DateTime)
                cell.DataType = ClosedXML.Excel.XLDataType.DateTime;
        }
    }
}
