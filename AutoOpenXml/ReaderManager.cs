using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class ReaderManager<T> where T : new()
    {
        internal IXLWorksheet ActiveWorksheet { get; set; }
        private int CurrentRowIndex { get; set; } = 1;

        internal void NextRow()
        {
            CurrentRowIndex++;
        }

        internal string TryReadColumnInfo(int index)
        {
            var value = ActiveWorksheet.Cell(CurrentRowIndex, index).Value ?? "";
            return value.ToString();
        }
    }
}
