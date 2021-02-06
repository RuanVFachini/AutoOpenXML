using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoOpenXml
{
    public class BaseWritterManager<T> : BaseDataProcessor<T> where T : new()
    {
        private List<string> Columns { get; set; }
        private int CurrentRowIndex { get; set; } = 1;

        internal void WriteFile()
        {
            Columns = DataToExport.First().Select(x => x.Key).ToList();
            WriteHeaders();
            WriteContentLines();
        }

        internal void WriteHeaders()
        {
            foreach (var column in Columns)
            {
                SetCellValue(column, column);
            }
            CurrentRowIndex++;
        }

        internal void WriteContentLines()
        {
            foreach (var rowData in DataToExport)
            {
                foreach (var column in Columns)
                {
                    rowData.TryGetValue(column, out var value);
                    SetCellValue(column, value);
                }
                CurrentRowIndex++;
            }
        }

        private int GetColumnIndex(string column)
        {
            return Columns.FindIndex(x => x == column) + 1;
        }
        private void SetCellValue(string column, string value)
        {
            var columnIndex = GetColumnIndex(column);
            ActiveWorksheet.Cell(CurrentRowIndex, columnIndex).Value = value;
        }
    }
}
