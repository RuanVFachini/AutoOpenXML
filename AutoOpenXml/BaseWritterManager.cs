using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

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
        private void SetCellValue(int columnIndex, Object value, ColumnTypes type = ColumnTypes.Text)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, columnIndex);

            if (type == ColumnTypes.Text)
            {
                cell.DataType = XLDataType.Text;
                cell.Value = (string) value;
            }
                

            if (type == ColumnTypes.Int)
            {
                cell.Value = Double.Parse(value.ToString());
                cell.DataType = XLDataType.Number;
            }

            if (type == ColumnTypes.Decimal)
            {
                cell.Value = Double.Parse(value.ToString());
                cell.DataType = XLDataType.Number;
            }


            if (type == ColumnTypes.DateTime)
            {
                cell.Value = ((DateTime) value);
                cell.DataType = XLDataType.DateTime;
                //cell.Style.DateFormat.Format = "dd/mm/yyyy hh:mm:ss";
            }

        }
    }
}
